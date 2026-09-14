; XRayXL -- the shared thunk for a patched VBA dispatch slot.
;
; HOW A SLOT IS ENTERED. The interpreter dispatches with
; `jmp qword ptr [rbx+rax*8]`, so a handler is entered by JMP: there is no
; return address, and every register the interpreter is using is live.
;
; WHY THE STUB CALLS US INSTEAD OF JUMPING.
; The obvious shape -- push the original, jump here, `ret` to it -- is exactly
; the return-address forgery that CET shadow stacks fault on, and a shadow-stack
; violation is STATUS_STACK_BUFFER_OVERRUN: a fast-fail that SEH cannot catch,
; so no circuit breaker would ever see it. `EXCEL.EXE` has no CET today, but the
; XLL side already refused to depend on that being true forever,
; and so does this. Instead the per-slot stub does:
;
;       call  XRayVbaBosThunk          <- real call, shadow stack balanced
;       jmp   qword ptr [rip+original]  <- tail jump, rsp exactly as it was
;
; so we return normally and the stub performs the tail jump. Nothing forges a
; return address, and the interpreter sees the original rsp on entry to the
; real handler.
;
; WHAT WE SAVE. Everything -- the general-purpose registers, RFLAGS,
; and (see SAVE_FP below) xmm0-xmm5 with MXCSR. The ABI argument that the
; interpreter keeps its live state in non-volatile registers is probably true --
; it must, because it calls C helpers -- but "probably" is not the standard for
; code on the path of every VBA statement in the process. Volatiles are cheap
; here relative to the call.
;
; FLAGS are saved too: a dispatch site is a computed branch and we cannot prove
; the flags are dead across it.
;
; ALIGNMENT. The interpreter's rsp at a dispatch site has no guaranteed
; alignment, so rsp is aligned explicitly before the call and restored from rbx
; afterwards rather than by unwinding the pushes.
;
; THE SAVED REGISTER BLOCK. rbx is passed to the handler as a second argument,
; because everything the interpreter was holding is now on the stack below it
; and that is where the interesting state lives -- r14 is the VBA frame base,
; and a function's arguments and its return value are reached through it.
; Offsets down from rbx, in push order. FIFTEEN pushes follow the one that
; establishes rbx, so the last slot is -0x78:
;
;   [rbx-08] rax   [rbx-30] r10   [rbx-58] r12
;   [rbx-10] rcx   [rbx-38] r11   [rbx-60] r13
;   [rbx-18] rdx   [rbx-40] rbp   [rbx-68] r14
;   [rbx-20] r8    [rbx-48] rsi   [rbx-70] r15
;   [rbx-28] r9    [rbx-50] rdi   [rbx-78] flags
;
; An earlier version of this table LEFT OUT r15 and so named every slot from
; -0x40 down one register too early. Nothing caught it: -0x70 is perfectly
; readable memory, so the guarded read succeeded and handed back r15 as though
; it were the frame base. It surfaced only because a dump printed the value and
; a frame base of 0xFFFFFFFFFFFFFFFE is obviously not one. Keep this table and
; the pushes below in the same field of view, and count them when either
; changes -- a comment is the only thing that ties this ABI together, and a
; wrong one here is invisible.

; UNWIND INFO, AND THE INVARIANT THAT MAKES IT SAFE.
;
; Both PROCs below are FRAME with .endprolog placed BEFORE the pushes, so the
; unwind info describes an EMPTY prologue. Nothing may therefore unwind
; through these frames: an unwinder would compute RSP as though the fifteen
; pushes (and the alignment and the xmm frame) had not happened, read a
; garbage return address, and jump to it.
;
; What holds that line is one line of C++: HookFilter in vbatrace.cpp returns
; EXCEPTION_EXECUTE_HANDLER for EVERY exception code, so an exception raised
; inside the hook is always handled inside XRayVbaOnStatement / XRayVbaOnExit and
; never propagates back out to here.
;
; If that ever changes, this asm must be fixed FIRST -- a .pushreg for each
; push and a .setframe rbx, 0 after `mov rbx, rsp`, with .endprolog moved
; below them -- because an escaping exception here produces an access
; violation with "faulting module unknown", which is the least debuggable
; failure this codebase can produce and has already cost one unexplained
; Excel. The two files are one mechanism; keep them in the same change.
;
; HookFilter cannot catch a stack overflow in the saves before the first call;
; the fixed 264-byte frame keeps that within the interpreter's own margin.

option casemap:none

EXTERN XRayVbaOnStatement : PROC
EXTERN XRayVbaOnExit      : PROC
EXTERN XRayVbaOnRaise     : PROC
EXTERN XRayVbaOnEnd       : PROC

.code

; 15 pushes after the one that establishes rbx.
SAVED_BYTES equ 120

; --------------------------------------------------------------------------
SAVE_ALL macro
    push    rbx
    mov     rbx, rsp            ; [rbx]=saved rbx  [rbx+8]=return addr  rbx+16=dispatch rsp
    push    rax
    push    rcx
    push    rdx
    push    r8
    push    r9
    push    r10
    push    r11
    push    rbp
    push    rsi
    push    rdi
    push    r12
    push    r13
    push    r14
    push    r15
    pushfq
endm

RESTORE_ALL macro
    lea     rsp, [rbx - SAVED_BYTES]
    popfq
    pop     r15
    pop     r14
    pop     r13
    pop     r12
    pop     rdi
    pop     rsi
    pop     rbp
    pop     r11
    pop     r10
    pop     r9
    pop     r8
    pop     rdx
    pop     rcx
    pop     rax
    pop     rbx
endm


; --------------------------------------------------------------------------
; FLOATING-POINT STATE. Saved for the same reason the general-purpose
; registers are: the interpreter reaches a handler by JMP, not by CALL, so
; nothing on its side spilled anything and every register is live.
;
; The x64 ABI's "xmm0-xmm5 are volatile" rule does NOT get us off this hook --
; that rule only binds a CALLER, and the interpreter never called us. Meanwhile
; the C++ hook definitely clobbers them: passing a double to _snprintf_s puts
; it in xmm0 by the ABI, and memcpy uses xmm for anything sizeable. So a live
; Double in the interpreter's xmm0 across a beginning-of-statement dispatch
; would come back holding our formatting intermediate, and the workbook would
; compute a WRONG NUMBER with no fault and no counter -- the one failure this
; project ranks worst.
;
; xmm6-xmm15 are non-volatile and the C++ callee preserves them, so six
; registers plus MXCSR is the whole exposure. MXCSR matters because the CRT's
; float formatting can change rounding and mask bits.
;
; Layout above rsp, which is 16-aligned here so movaps is safe:
;   [rsp+00..1F]  shadow space for the callee (must be first)
;   [rsp+20..7F]  xmm0..xmm5
;   [rsp+80..83]  mxcsr
XMM_FRAME equ 144           ; 32 shadow + 96 xmm + 16 mxcsr/pad, 16-aligned

SAVE_FP macro
    sub     rsp, XMM_FRAME
    movaps  xmmword ptr [rsp+20h], xmm0
    movaps  xmmword ptr [rsp+30h], xmm1
    movaps  xmmword ptr [rsp+40h], xmm2
    movaps  xmmword ptr [rsp+50h], xmm3
    movaps  xmmword ptr [rsp+60h], xmm4
    movaps  xmmword ptr [rsp+70h], xmm5
    stmxcsr dword ptr [rsp+80h]
endm

RESTORE_FP macro
    ldmxcsr dword ptr [rsp+80h]
    movaps  xmm0, xmmword ptr [rsp+20h]
    movaps  xmm1, xmmword ptr [rsp+30h]
    movaps  xmm2, xmmword ptr [rsp+40h]
    movaps  xmm3, xmmword ptr [rsp+50h]
    movaps  xmm4, xmmword ptr [rsp+60h]
    movaps  xmm5, xmmword ptr [rsp+70h]
    add     rsp, XMM_FRAME
endm

; --------------------------------------------------------------------------
; void XRayVbaBosThunk(void)   -- entered by CALL from a per-slot stub.
; One shape for all four hooks: save everything, call the recorder, restore.
HOOK_THUNK MACRO thunkName, recorder
thunkName PROC FRAME
    .endprolog
    SAVE_ALL
    lea     rcx, [rbx + 16]     ; rcx = the interpreter's rsp at the dispatch
    mov     rdx, rbx            ; rdx = top of the saved register block
    and     rsp, -16
    SAVE_FP                     ; xmm0-5 + mxcsr, and the callee's shadow space
    call    recorder
    RESTORE_FP
    RESTORE_ALL
    ret
thunkName ENDP
ENDM

HOOK_THUNK XRayVbaBosThunk,   XRayVbaOnStatement
HOOK_THUNK XRayVbaExitThunk,  XRayVbaOnExit
; The raise opcode (slot 497) fires four times per Err.Raise; the recorder dedupes.
HOOK_THUNK XRayVbaRaiseThunk, XRayVbaOnRaise
; The End opcode (slot 619) fires no exit for the frames it kills, and runs before its handler.
HOOK_THUNK XRayVbaEndThunk,   XRayVbaOnEnd

END
