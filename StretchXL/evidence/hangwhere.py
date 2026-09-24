"""WHY WILL IT NOT EXIT? -- every thread in a hang dump, and what it is parked on.

dumpstack.py answers "where did it die", which needs an exception record and
a faulting thread. A hang has neither. It has something
better: every thread is still sitting exactly where it stuck, with nothing
unwound and no combase handler standing between the cause and the report.

So this walks ALL threads. For each it reads the saved CONTEXT for the parked
Rip, then scans the thread's own stack for values that land in a module's
executable range -- a return-address sieve, not a real unwind, so it over-
reports; a name here means "this module appears on that stack", not "this is a
live frame". That is still decisive for the question being asked, which is
whether the threads holding the process open have anything to do with us.

    python hangwhere.py <dump> [--all] [--ours=name1,name2]

--ours names the modules of interest (substring match, case-insensitive) --
in StretchXL a suite declares them in its suite.psd1 and the manager passes
them through. With no --ours, no module is special and every non-idle thread
is shown. Without --all it prints the threads that matter: the ones
referencing a module of interest, plus anything not parked in a plain wait.
"""
import os, sys, mmap, struct, collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dumpstack import streams, modules, threads, memranges, readmem, owner, exception

# x64 CONTEXT: Rsp at 0x98, Rbp at 0xA0, Rip at 0xF8.
CTX_RSP, CTX_RBP, CTX_RIP = 0x98, 0xA0, 0xF8

# from --ours, since StretchXL knows no product's name; empty means none is special
OURS = ()

# A thread parked in one of these is idle by design -- a worker waiting for
# work, not a thread holding the process open. They are counted, not dropped.
IDLE = ('ntdll.dll', 'win32u.dll', 'kernelbase.dll', 'ntoskrnl.exe')


def stackrefs(mm, rngs, mods, rsp, start, size):
    """Distinct modules referenced from [rsp, stack top), in stack order."""
    top = start + size
    if not (start <= rsp < top):
        rsp = start                      # context unavailable; scan it all
    data = readmem(rngs, mm, rsp, min(top - rsp, 512 * 1024))
    seen, order = set(), []
    for o in range(0, len(data) - 8, 8):
        v = struct.unpack_from('<Q', data, o)[0]
        if v < 0x10000:
            continue
        name, _ = owner(mods, v)
        if name and name not in seen:
            seen.add(name)
            order.append(name)
    return order


def main(path, show_all=False, ours=()):
    global OURS
    OURS = tuple(k.lower() for k in ours if k)
    with open(path, 'rb') as f:
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
    st = streams(mm)
    mods, thrs, rngs = modules(mm, st), threads(mm, st), memranges(mm, st)

    # A crash dump's faulting thread is never filtered out: a thread being
    # killed parks in ntdll, which the idle rule would otherwise hide.
    exc = exception(mm, st)

    print('%s\n  %d module(s), %d thread(s)' % (os.path.basename(path), len(mods), len(thrs)))
    for base, size, name in mods:
        if any(k in name.lower() for k in OURS):
            print('  loaded: %-22s base %016x  size %x' % (name, base, size))

    if exc:
        nm, rva = owner(mods, exc['addr'])
        print("  EXCEPTION on tid %d: code 0x%08X at %016x (%s+0x%x)"
              % (exc['tid'], exc['code'], exc['addr'], nm or 'NOT IN ANY MODULE', rva))
        if exc['code'] == 0xC0000409 and exc['params']:
            extra = ' = GUARD_ICALL_CHECK_FAILURE (CFG rejected an indirect call)' if exc['params'][0] == 10 else ''
            print("    __fastfail code %d%s" % (exc['params'][0], extra))
        elif exc['code'] == 0xC0000005 and len(exc['params']) >= 2:
            kind = {0: 'read', 1: 'write', 8: 'execute'}.get(exc['params'][0], '?')
            print("    access violation: %s of %016x" % (kind, exc['params'][1]))

    every, interesting, idle, tally = [], [], 0, collections.Counter()
    parkedHist = collections.Counter()
    for ti, t in enumerate(thrs):
        rip = rsp = 0
        if t['ctxSize'] >= 0x100:
            rip = struct.unpack_from('<Q', mm, t['ctxRva'] + CTX_RIP)[0]
            rsp = struct.unpack_from('<Q', mm, t['ctxRva'] + CTX_RSP)[0]
        parked, prva = owner(mods, rip)
        refs = stackrefs(mm, rngs, mods, rsp, t['start'], t['size'])
        mine = [r for r in refs if any(k in r.lower() for k in OURS)]
        for r in refs:
            tally[r] += 1
        parkedHist['%s+0x%x' % (parked or '?', prva)] += 1
        rec = dict(tid=t['tid'], idx=ti, rip=rip, parked=parked or '?', rva=prva, refs=refs, mine=mine)
        every.append(rec)
        # ti == 0 is the main thread: never filtered, for the reason above.
        if ti == 0 or (exc and t['tid'] == exc['tid']) or mine or (parked and parked.lower() not in IDLE):
            rec['keep'] = True
            interesting.append(rec)
        else:
            rec['keep'] = False
            idle += 1

    print('\n  %d thread(s) parked in a plain wait and referencing nothing of ours' % idle)
    # --all includes the idle threads, marked, so the filter can be checked
    shown = every if show_all else interesting
    if show_all:
        print('  all %d thread(s), the idle ones marked:\n' % len(shown))
    else:
        print('  %d thread(s) worth looking at:\n' % len(shown))
    for r in shown:
        flag = '  <== OURS' if r['mine'] else ''
        if exc and r['tid'] == exc['tid']:
            flag += '  <== FAULTING THREAD'
        if r['idx'] == 0:
            flag += '  <== MAIN THREAD'
        if not r['keep']:
            flag += '  (idle)'
        print('  tid %-6d parked in %s+0x%x%s' % (r['tid'], r['parked'], r['rva'], flag))
        print('      stack references: %s' % (', '.join(r['refs'][:12]) or '(none resolved)'))
        if r['mine']:
            print('      OURS ON THIS STACK: %s' % ', '.join(r['mine']))
    if not shown:
        label = '/'.join(OURS) if OURS else 'any module of interest'
        print('  (none -- every thread is idle and none references %s)' % label)

    print('')
    print('  where every thread is parked (%d distinct):' % len(parkedHist))
    for loc, c in parkedHist.most_common(10):
        print('    %-42s %d thread(s)' % (loc, c))

    print('')
    print('  VERDICT')
    if not OURS:
        print('    No modules of interest were named (--ours=...); the non-idle threads')
        print('    above are the evidence. Name the modules you suspect for a verdict.')
        return
    hits = {m: c for m, c in tally.items() if any(k in m.lower() for k in OURS)}
    label = '/'.join(OURS)
    if hits:
        print('    %s appear on %s' % (label, ', '.join('%s: %d stack(s)' % (m, c) for m, c in hits.items())))
        print('    That is a reference on a stack, not proof of a live frame -- but it is')
        print('    the opposite of absence, and absence is what would exonerate them.')
    else:
        print('    None of [%s] appears on ANY thread stack in this dump.' % label)
        print('    Note what that does and does not say: it is evidence, not an alibi,')
        print('    because a thread can hold the process open through a reference it')
        print('    handed out and never appear on a stack at all.')


if __name__ == '__main__':
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    if not a:
        print(__doc__)
        sys.exit(2)
    ours_arg = ()
    for x in sys.argv[1:]:
        if x.startswith('--ours='):
            ours_arg = tuple(p.strip() for p in x[len('--ours='):].split(',') if p.strip())
    main(a[0], '--all' in sys.argv, ours_arg)
