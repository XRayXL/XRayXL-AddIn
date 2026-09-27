// Converts XRayXL traces with the page's own converter, decodes the protobuf it
// wrote, and checks it against the trace read by a separate parser.
//   node perfetto/check.mjs [--out <dir>] <trace.csv|trace.jsonl>...
import { readFileSync, writeFileSync, mkdirSync, createReadStream } from 'node:fs';
import { basename, join } from 'node:path';

const html = readFileSync(new URL('./XRayXL-Perfetto.html', import.meta.url), 'utf8');
const script = /<script id="converter">([\s\S]*?)<\/script>/.exec(html);
if (!script) throw new Error('no converter script in XRayXL-Perfetto.html');
new Function(script[1])();
const { Converter, HEADER } = globalThis.XRayXLPerfetto;

async function convert(path) {
  const conv = new Converter(basename(path));
  const dec = new TextDecoder('windows-1252');
  for await (const chunk of createReadStream(path, { highWaterMark: 1 << 16 })) {
    conv.push(dec.decode(chunk, { stream: true }));
  }
  conv.push(dec.decode());
  return conv.finish();
}

// ---- the trace, read independently of the converter ----
function readTrace(path) {
  const lines = readFileSync(path, 'latin1').split('\r\n');
  if (lines[lines.length - 1] === '') lines.pop();
  return lines[0].startsWith('{') ? readJsonl(lines) : readCsv(lines);
}

function readJsonl(lines) {
  let frequency = 10000000n;
  const entries = new Map(), exits = new Map(), events = new Map(), inputs = [];
  let depthCapped = 0;
  for (const line of lines) {
    const o = JSON.parse(line);
    if (o.kind === 'event' && o.function === 'arm') {
      const hz = (o.args || []).find((a) => a.name === 'qpcFrequency');
      if (hz) frequency = BigInt(hz.value.v);
    }
    inputs.push(Number(o.input));
    const row = { q: BigInt(o.qpc), thread: String(o.thread), trust: o.trust ?? '', ticks: String(o.ticks ?? ''),
      fn: o.function ?? '' };
    if (o.kind === 'entry') entries.set(String(o.span), row);
    else if (o.kind === 'exit') exits.set(String(o.span), row);
    else if (o.kind === 'event') events.set(String(o.seq), row);
    else if (o.kind === 'depth-capped') depthCapped++;
  }
  return summarise(entries, exits, events, depthCapped, inputs, frequency);
}

function summarise(entries, exits, events, depthCapped, inputs, frequency) {
  inputs.sort((a, b) => a - b);
  const dropped = inputs.length ? inputs[inputs.length - 1] - inputs.length : 0;
  return { entries, exits, events, depthCapped, dropped, ns: (q) => q * 1000000000n / frequency };
}

function readCsv(lines) {
  // The breaks column is optional and last, so every other column keeps its place.
  const breaks = lines[0] === HEADER + ',breaks';
  if (lines[0] !== HEADER && !breaks) throw new Error('header differs');
  const cols = HEADER.split(',');
  const width = cols.length + (breaks ? 1 : 0);
  let frequency = 10000000n;
  const col = Object.fromEntries(cols.map((n, i) => [n, i]));
  const field = /("(?:[^"]|"")*"|[^,]*)(,|$)/g;
  const entries = new Map(), exits = new Map(), events = new Map(), inputs = [];
  let depthCapped = 0;
  for (let i = 1; i < lines.length; i++) {
    const f = [];
    field.lastIndex = 0;
    for (let m; (m = field.exec(lines[i])); ) {
      f.push(m[1].startsWith('"') ? m[1].slice(1, -1).replace(/""/g, '"') : m[1]);
      if (m[2] === '') break;
    }
    if (f.length !== width) throw new Error(`line ${i + 1}: ${f.length} fields`);
    const hz = f[col.kind] === 'event' && f[col.function] === 'arm' && /(?:^| )qpcFrequency=(\d+)/.exec(f[col.args]);
    if (hz) frequency = BigInt(hz[1]);
    inputs.push(Number(f[col.input]));
    const row = { q: BigInt(f[col.qpc]), thread: f[col.thread], trust: f[col.trust], ticks: f[col.ticks], fn: f[col.function] };
    if (f[col.kind] === 'entry') entries.set(f[col.span], row);
    else if (f[col.kind] === 'exit') exits.set(f[col.span], row);
    else if (f[col.kind] === 'event') events.set(f[col.seq], row);
    else if (f[col.kind] === 'depth-capped') depthCapped++;
  }
  return summarise(entries, exits, events, depthCapped, inputs, frequency);
}

// ---- protobuf decoding ----
function varint(buf, pos) {
  let v = 0, mul = 1, b;
  do { b = buf[pos.i++]; v += (b & 0x7f) * mul; mul *= 128; } while (b & 0x80);
  return v;
}
function varintBig(buf, pos) {
  let v = 0n, shift = 0n, b;
  do { b = buf[pos.i++]; v |= BigInt(b & 0x7f) << shift; shift += 7n; } while (b & 0x80);
  return v;
}
function* fields(buf, start, end) {
  const pos = { i: start };
  while (pos.i < end) {
    const key = varint(buf, pos);
    const f = Math.floor(key / 8), wt = key % 8;
    if (wt === 0) {
      const at = pos.i;
      yield { f, wt, v: varint(buf, pos), at };
    } else if (wt === 2) {
      const len = varint(buf, pos);
      yield { f, wt, s: pos.i, e: pos.i + len };
      pos.i += len;
    } else {
      throw new Error(`unexpected wire type ${wt} at ${pos.i}`);
    }
  }
}
const text = (buf, fld) => new TextDecoder().decode(buf.subarray(fld.s, fld.e));

function check(path, result) {
  const buf = Buffer.concat(result.parts.map((p) => Buffer.from(p.buffer, p.byteOffset, p.byteLength)));
  const csv = readTrace(path);
  const fail = [];
  const names = new Map(), annNames = new Map(), tracks = new Map(), stacks = new Map();
  let packets = 0, prevTs = -1n, instants = 0, asyncInstants = 0, capped = 0, droppedMarked = 0, eventMarks = 0;
  let pairs = 0, reconstructed = 0, ticksDiffer = 0;

  for (const top of fields(buf, 0, buf.length)) {
    if (top.f !== 1 || top.wt !== 2) { fail.push('top-level field is not a TracePacket'); break; }
    const pkt = { flags: 0, seq: 0, ts: null, ev: null, td: null, interned: null };
    for (const fl of fields(buf, top.s, top.e)) {
      if (fl.f === 8) pkt.ts = varintBig(buf, { i: fl.at });
      else if (fl.f === 10) pkt.seq = fl.v;
      else if (fl.f === 13) pkt.flags = fl.v;
      else if (fl.f === 11) pkt.ev = fl;
      else if (fl.f === 60) pkt.td = fl;
      else if (fl.f === 12) pkt.interned = fl;
    }
    if (packets === 0 && (pkt.flags !== 1 || !pkt.interned)) fail.push('first packet does not clear state and intern names');
    packets++;
    if (pkt.seq !== 1) fail.push(`packet ${packets}: sequence id ${pkt.seq}`);
    if (pkt.interned) {
      for (const fl of fields(buf, pkt.interned.s, pkt.interned.e)) {
        let iid = 0, name = '';
        for (const g of fields(buf, fl.s, fl.e)) { if (g.f === 1) iid = g.v; else if (g.f === 2) name = text(buf, g); }
        if (fl.f === 2) names.set(iid, name);
        else if (fl.f === 3) annNames.set(iid, name);
      }
    }
    if (pkt.td) {
      const t = { name: '', tid: null, parent: 0 };
      let uuid = 0;
      for (const fl of fields(buf, pkt.td.s, pkt.td.e)) {
        if (fl.f === 1) uuid = fl.v;
        else if (fl.f === 2) t.name = text(buf, fl);
        else if (fl.f === 5) t.parent = fl.v;
        else if (fl.f === 4) for (const g of fields(buf, fl.s, fl.e)) if (g.f === 2) t.tid = String(g.v);
      }
      tracks.set(uuid, t);
    }
    if (!pkt.ev) continue;

    if (pkt.flags !== 2) fail.push(`packet ${packets}: event without SEQ_NEEDS_INCREMENTAL_STATE`);
    if (pkt.ts < prevTs) fail.push(`packet ${packets}: timestamp goes backwards`);
    prevTs = pkt.ts;
    let type = 0, uuid = 0, nameIid = 0;
    const ann = {};
    for (const fl of fields(buf, pkt.ev.s, pkt.ev.e)) {
      if (fl.f === 9) type = fl.v;
      else if (fl.f === 11) uuid = fl.v;
      else if (fl.f === 10) nameIid = fl.v;
      else if (fl.f === 4) {
        let key = '', value;
        for (const g of fields(buf, fl.s, fl.e)) {
          if (g.f === 1) key = annNames.get(g.v);
          else if (g.f === 4) value = g.v;
          else if (g.f === 6) value = text(buf, g);
        }
        ann[key] = value;
      }
    }
    const track = tracks.get(uuid);
    if (!track) { fail.push(`packet ${packets}: track ${uuid} has no descriptor before it`); continue; }
    const tid = track.tid ?? tracks.get(track.parent)?.tid ?? null;
    if (type !== 2 && !names.has(nameIid)) fail.push(`packet ${packets}: name iid ${nameIid} not interned`);
    if (!stacks.has(uuid)) stacks.set(uuid, []);
    const stack = stacks.get(uuid);

    if (type === 1) {
      const span = String(ann.span);
      const entry = csv.entries.get(span), exit = csv.exits.get(span);
      if (ann['entry seq'] !== undefined) {
        if (!entry) fail.push(`span ${span}: slice with no entry row`);
        else {
          if (pkt.ts !== csv.ns(entry.q)) fail.push(`span ${span}: begins at ${pkt.ts}, entry qpc ${entry.q}`);
          if (tid !== entry.thread) fail.push(`span ${span}: on thread ${tid}, entry thread ${entry.thread}`);
          if (/^[\x20-\x7e]*$/.test(entry.fn) && entry.fn && names.get(nameIid) !== entry.fn) fail.push(`span ${span}: named ${names.get(nameIid)}`);
        }
      } else {
        reconstructed++;
        if (!exit || entry) fail.push(`span ${span}: reconstructed slice, but the CSV has entry ${!!entry} exit ${!!exit}`);
        else if (pkt.ts !== csv.ns(exit.q - BigInt(exit.ticks))) fail.push(`span ${span}: reconstructed start is wrong`);
      }
      stack.push({ span, entry, exit });
    } else if (type === 2) {
      const open = stack.pop();
      if (!open) { fail.push(`packet ${packets}: slice end on track ${uuid} with nothing open`); continue; }
      if (!open.exit) { fail.push(`span ${open.span}: closed, but the CSV has no exit`); continue; }
      const expected = csv.ns(open.entry && open.exit.q < open.entry.q ? open.entry.q : open.exit.q);
      if (pkt.ts !== expected) fail.push(`span ${open.span}: ends at ${pkt.ts}, expected ${expected} (nesting is wrong)`);
      if (open.entry) {
        pairs++;
        if (/^\d+$/.test(open.exit.ticks) && BigInt(open.exit.ticks) !== open.exit.q - open.entry.q) ticksDiffer++;
      }
    } else if (type === 3) {
      instants++;
      if (ann.trust === 'async') asyncInstants++;
      if (names.get(nameIid) === 'depth-capped') capped++;
      else if (ann.seq !== undefined) {
        const ev = csv.events.get(String(ann.seq));
        eventMarks++;
        if (!ev) fail.push(`packet ${packets}: marker for seq ${ann.seq}, which is not an event row`);
        else {
          if (pkt.ts !== csv.ns(ev.q)) fail.push(`event seq ${ann.seq}: at ${pkt.ts}, its row's qpc ${ev.q}`);
          if (tid !== ev.thread) fail.push(`event seq ${ann.seq}: on thread ${tid}, its row's thread ${ev.thread}`);
          if (names.get(nameIid) !== ev.fn) fail.push(`event seq ${ann.seq}: named ${names.get(nameIid)}, not ${ev.fn}`);
        }
      }
      if (typeof ann['rows dropped'] === 'number') droppedMarked += ann['rows dropped'];
    } else {
      fail.push(`packet ${packets}: event type ${type}`);
    }
  }

  let stillOpen = 0;
  for (const stack of stacks.values()) for (const s of stack) {
    stillOpen++;
    if (!s.entry || s.exit) fail.push(`span ${s.span}: left open, but the CSV has entry ${!!s.entry} exit ${!!s.exit}`);
  }
  let expectedPairs = 0, expectedOpen = 0, expectedAsync = 0;
  for (const [span, e] of csv.entries) {
    const x = csv.exits.get(span);
    if (!x) expectedOpen++;
    else if (x.trust === 'async') expectedAsync++;
    else expectedPairs++;
  }
  if (pairs !== expectedPairs) fail.push(`${pairs} slices closed, the CSV pairs ${expectedPairs}`);
  if (stillOpen !== expectedOpen) fail.push(`${stillOpen} slices open, the CSV has ${expectedOpen} entries without exits`);
  if (asyncInstants !== expectedAsync) fail.push(`${asyncInstants} async markers, the CSV has ${expectedAsync}`);
  if (eventMarks !== csv.events.size) fail.push(`${eventMarks} event markers, the trace has ${csv.events.size} event rows`);
  if (capped !== csv.depthCapped) fail.push(`${capped} depth-capped markers, the CSV has ${csv.depthCapped}`);
  if (droppedMarked !== csv.dropped) fail.push(`${droppedMarked} rows marked dropped, the CSV is missing ${csv.dropped}`);

  const extra = [...tracks.values()].filter((t) => t.parent);
  return {
    fail, packets, pairs, stillOpen, reconstructed, instants, ticksDiffer,
    overlapTracks: extra.filter((t) => t.name.includes('overlapping')).length,
    upperTracks: extra.filter((t) => t.name.includes('upper-bound')).length,
  };
}

const args = process.argv.slice(2);
let outDir = null;
const files = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--out') outDir = args[++i];
  else files.push(args[i]);
}
if (!files.length) {
  console.error('usage: node perfetto/check.mjs [--out <dir>] <trace.csv|trace.jsonl>...');
  process.exit(2);
}

let failed = 0;
for (const path of files) {
  const t0 = performance.now();
  let result;
  try {
    result = await convert(path);
  } catch (err) {
    failed++;
    console.log(`FAIL ${path}\n  conversion threw: ${err.message}`);
    continue;
  }
  const ms = performance.now() - t0;
  if (outDir) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(join(outDir, basename(path).replace(/\.(?:csv|jsonl)$/i, '') + '.pftrace'), Buffer.concat(result.parts));
  }
  const c = check(path, result);
  const s = result.stats;
  const verdict = c.fail.length ? 'FAIL' : 'PASS';
  if (c.fail.length) failed++;
  console.log(`${verdict} ${basename(path)}  ${s.rows} rows -> ${c.packets} packets, ${(result.bytes / 1e6).toFixed(2)} MB in ${(ms / 1000).toFixed(2)} s`);
  console.log(`  calls ${s.calls} (closed ${c.pairs}, open ${c.stillOpen}, reconstructed ${c.reconstructed}), instants ${c.instants} (events ${s.events}), ` +
    `threads ${s.threads}, overlapping tracks ${c.overlapTracks}, upper-bound tracks ${c.upperTracks}, dropped ${s.dropped}, ` +
    `ticks != exit-entry qpc on ${c.ticksDiffer}`);
  for (const f of c.fail.slice(0, 10)) console.log('  - ' + f);
  if (c.fail.length > 10) console.log(`  ... and ${c.fail.length - 10} more`);
}
process.exit(failed ? 1 : 0);
