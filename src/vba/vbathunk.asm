; XRayXL -- the shared thunk for a patched VBA dispatch slot.
;
; HOW A SLOT IS ENTERED. The interpreter dispatches with
; `jmp qword ptr [rbx+rax*8]`, so a handler is entered by JMP: there is no
; return address, and every register the interpreter is using is live.
;
; The stub calls here instead of jumping: push-the-original-and-`ret` forges a return address,
; which CET shadow stacks fast-fail on, past any SEH. So the per-slot stub does:
;
;       call  XRayVbaBosThunk          <- real call, shadow stack balanced
;       jmp   qword ptr [rip+original]  <- tail jump, rsp exactly as it was
;
; We return normally and the stub performs the tail jump, so the interpreter sees its original
; rsp on entry to the real handler.
;
; Everything is saved: the general-purpose registers, RFLAGS, and (SAVE_FP below) xmm0-xmm5 with
; MXCSR. The interpreter jumped here, so no register can be assumed dead.
;
; FLAGS are saved too: a dispatch site is a computed branch and we cannot prove
; the flags are dead across it.
;
; ALIGNMENT. The interpreter's rsp at a dispatch site has no guaranteed
; alignment, so rsp is aligned explicitly before the call and restored from rbx
; afterwards rather than by unwinding the pushes.
;
; rbx is passed to the handler as a second argument: the interpreter's state is on the stack
; below it, and r14 there is the VBA frame base. Offsets down from rbx, in push order; fifteen
; pushes follow the one that establishes rbx, so the last slot is -0x78:
;
;   [rbx-08] rax   [rbx-30] r10   [rbx-58] r12
;   [rbx-10] rcx   [rbx-38] r11   [rbx-60] r13
;   [rbx-18] rdx   [rbx-40] rbp   [rbx-68] r14
;   [rbx-20] r8    [rbx-48] rsi   [rbx-70] r15
;   [rbx-28] r9    [rbx-50] rdi   [rbx-78] flags
;
; Keep this table and the pushes below in step, and count them when either changes: a wrong
; offset still reads valid memory, just the wrong register.

; Unwind info. Both PROCs are FRAME with .endprolog before the pushes, so the unwind info
; describes an empty prologue and nothing may unwind through these frames. HookFilter in
; vbatrace.cpp guarantees that by returning EXCEPTION_EXECUTE_HANDLER for every exception code.
;
; If that ever changes, add a .pushreg for each push and a .setframe rbx, 0 after `mov rbx,
; rsp`, with .endprolog below them, in the same change.
;
; HookFilter cannot catch a stack overflow in the saves before the first call; the fixed
; 264-byte frame keeps that within the interpreter's own margin.

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
; Floating-point state. The interpreter reaches a handler by JMP, so the ABI's volatile-register
; rule does not apply, and the C++ hook does clobber xmm0-xmm5 (a double passed to _snprintf_s,
; memcpy). A live Double there would come back wrong with no fault. xmm6-xmm15 are preserved by
; the callee. MXCSR is saved because the CRT's float formatting can change rounding and mask
; bits.
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
