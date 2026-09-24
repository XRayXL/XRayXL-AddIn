# Whose code is on the faulting thread's stack? A crash inside a system module such
# as combase is often only the messenger; the culprit is further down the stack.
#
# It does not unwind (x64 needs .pdata for every module); it scans the stack for
# 8-byte values inside a loaded module. A module's absence is the stronger signal,
# since a hit may be a coincidental value.
#
#     python dumpstack.py [<dump> ...] [--ours=name1,name2]
#
# --ours names the modules of interest (prefix, case-insensitive), because StretchXL
# knows no product's name.
import struct, sys, os, glob, mmap

# Modules of interest, from --ours. Empty means none is special.
OURS = ()

EXCEPTION_STREAM, THREAD_LIST, MODULE_LIST, MEMORY_LIST, MEMORY64_LIST = 6, 3, 4, 5, 9
CODES = {0xC0000005: 'ACCESS_VIOLATION',
         0xC0000409: 'STATUS_STACK_BUFFER_OVERRUN (__fastfail) -- SEH CANNOT CATCH THIS',
         0xC000001D: 'ILLEGAL_INSTRUCTION',
         0xC0000374: 'HEAP_CORRUPTION',
         0x80000003: 'BREAKPOINT'}


def streams(mm):
    """type -> (size, rva) for every stream in the dump directory."""
    _, _, n, dr = struct.unpack_from('<IIII', mm, 0)
    out = {}
    for i in range(n):
        t, size, rva = struct.unpack_from('<III', mm, dr + i * 12)
        out[t] = (size, rva)
    return out


def mdstring(mm, rva):
    n = struct.unpack_from('<I', mm, rva)[0]
    return mm[rva + 4: rva + 4 + n].decode('utf-16-le', 'replace')


def modules(mm, st):
    if MODULE_LIST not in st:
        return []
    _, rva = st[MODULE_LIST]
    n = struct.unpack_from('<I', mm, rva)[0]
    out = []
    for i in range(n):
        o = rva + 4 + i * 108
        base, size, _, _, nameRva = struct.unpack_from('<QIIII', mm, o)
        out.append((base, size, os.path.basename(mdstring(mm, nameRva))))
    return sorted(out)


def threads(mm, st):
    if THREAD_LIST not in st:
        return []
    _, rva = st[THREAD_LIST]
    n = struct.unpack_from('<I', mm, rva)[0]
    out = []
    for i in range(n):
        o = rva + 4 + i * 48
        tid, _, _, _, _teb, stkStart, stkSize, stkRva, ctxSize, ctxRva = \
            struct.unpack_from('<IIIIQQIIII', mm, o)
        out.append(dict(tid=tid, start=stkStart, size=stkSize, rva=stkRva,
                        ctxSize=ctxSize, ctxRva=ctxRva))
    return out


def exception(mm, st):
    if EXCEPTION_STREAM not in st:
        return None
    _, rva = st[EXCEPTION_STREAM]
    tid = struct.unpack_from('<I', mm, rva)[0]
    code, flags, _rec, addr, nparam = struct.unpack_from('<IIQQI', mm, rva + 8)
    params = [struct.unpack_from('<Q', mm, rva + 8 + 32 + k * 8)[0]
              for k in range(min(nparam, 15))]
    return dict(tid=tid, code=code, addr=addr, params=params)


def memranges(mm, st):
    """(virtual address, size, file offset) for every captured region.

    A full dump keeps all memory in Memory64List, where the per-thread stack
    descriptor's Rva does not point at the stack bytes, so addresses are
    resolved through this map rather than through that Rva.
    """
    out = []
    if MEMORY64_LIST in st:
        _, m = st[MEMORY64_LIST]
        cnt, base = struct.unpack_from('<QQ', mm, m)
        off = base
        for i in range(cnt):
            sa, sz = struct.unpack_from('<QQ', mm, m + 16 + i * 16)
            out.append((sa, sz, off)); off += sz
    elif MEMORY_LIST in st:
        _, m = st[MEMORY_LIST]
        c = struct.unpack_from('<I', mm, m)[0]
        for i in range(c):
            sa, dsz, drva = struct.unpack_from('<QII', mm, m + 4 + i * 16)
            out.append((sa, dsz, drva))
    return sorted(out)


def readmem(rngs, mm, va, n):
    for sa, sz, fo in rngs:
        if sa <= va < sa + sz:
            avail = min(n, (sa + sz) - va)
            o = fo + (va - sa)
            return mm[o:o + avail]
    return b''


def owner(mods, addr):
    for base, size, name in mods:
        if base <= addr < base + size:
            return name, addr - base
    return None, 0


def main(path):
    with open(path, 'rb') as f:
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
    st = streams(mm)
    mods = modules(mm, st)
    exc = exception(mm, st)

    print("dump   : %s (%.0f MB, %d modules)" % (os.path.basename(path),
                                                 os.path.getsize(path) / 1048576, len(mods)))
    if not exc:
        print("  no exception stream -- not a crash dump")
        return 2

    mname, moff = owner(mods, exc['addr'])
    print("  code   : 0x%08X  %s" % (exc['code'], CODES.get(exc['code'], '?')))
    print("  at     : 0x%016X  %s" % (exc['addr'],
                                      ("%s+0x%X" % (mname, moff)) if mname else "NO MODULE (private memory)"))
    if exc['code'] == 0xC0000005 and len(exc['params']) >= 2:
        kind = {0: 'read', 1: 'write', 8: 'execute'}.get(exc['params'][0], '?')
        print("  tried  : %s of 0x%016X" % (kind, exc['params'][1]))
    print("  thread : %d" % exc['tid'])

    th = [t for t in threads(mm, st) if t['tid'] == exc['tid']]
    if not th:
        print("  faulting thread not in the thread list")
        return 1
    t = th[0]
    rngs = memranges(mm, st)
    stack = readmem(rngs, mm, t['start'], t['size'])
    if not stack:
        print("  the thread's stack range is not in the captured memory")
        return 1
    print("  stack  : 0x%X..0x%X (%d KB captured)\n"
          % (t['start'], t['start'] + t['size'], t['size'] // 1024))

    # Every 8-byte aligned value on the stack that lands in a module.
    hits, seen = [], {}
    for off in range(0, len(stack) - 8, 8):
        v = struct.unpack_from('<Q', stack, off)[0]
        if v < 0x10000:
            continue
        name, mo = owner(mods, v)
        if name:
            hits.append((t['start'] + off, name, mo))
            seen[name] = seen.get(name, 0) + 1

    print("  modules referenced from this stack, most first:")
    for name, n in sorted(seen.items(), key=lambda kv: -kv[1]):
        print("    %-28s %5d" % (name, n))

    # with nothing named, say so rather than claim "not referenced" about an empty list
    print()
    if not OURS:
        print("  No modules of interest were named (--ours=...); the table above")
        print("  is the whole answer.")
    else:
        named = ', '.join(OURS)
        ours = [h for h in hits if h[1].lower().startswith(OURS)]
        if ours:
            print("  *** MATCH: %s on the faulting thread's stack, %d reference(s) ***"
                  % (named, len(ours)))
            for a, n, o in ours[:20]:
                print("      stack 0x%016X -> %s+0x%X" % (a, n, o))
        else:
            print("  No reference to %s anywhere on the faulting thread's stack." % named)
            print("  (That is evidence, not proof: damage done earlier, on another")
            print("   thread, would leave no trace here.)")

    # The frames nearest the fault, in stack order -- the closest this gets to
    # a call stack without unwinding.
    print("\n  first 24 module references from the top of the stack:")
    for a, n, o in hits[:24]:
        print("    0x%016X  %s+0x%X" % (a, n, o))
    return 0


if __name__ == '__main__':
    args = [x for x in sys.argv[1:] if not x.startswith('--')]
    for x in sys.argv[1:]:
        if x.startswith('--ours='):
            OURS = tuple(p.strip().lower() for p in x[len('--ours='):].split(',') if p.strip())
    if args:
        paths = args
    else:
        d = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'CrashDumps')
        paths = sorted(glob.glob(os.path.join(d, 'EXCEL*.dmp')), key=os.path.getmtime)[-1:]
    rc = 0
    for p in paths:
        rc |= main(p)
        print()
    sys.exit(rc)
