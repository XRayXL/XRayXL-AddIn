; XRayXL -- the shared thunk for a patched VBA dispatch slot.
;
; The interpreter enters a handler by JMP, so every register and flag is live and all are saved.
; The stub calls here and then tail-jumps to the original, because push-and-`ret` would forge a
; return address, which CET shadow stacks fast-fail on:
;
;       call  XRayVbaBosThunk          <- real call, shadow stack balanced
;       jmp   qword ptr [rip+original]  <- tail jump, rsp exactly as it was
;
; The dispatch rsp has no guaranteed alignment, so rsp is aligned before the call and restored
; from rbx. rbx goes to the handler as its second argument; offsets down from it, in push order:
;
;   [rbx-08] rax   [rbx-30] r10   [rbx-58] r12
;   [rbx-10] rcx   [rbx-38] r11   [rbx-60] r13
;   [rbx-18] rdx   [rbx-40] rbp   [rbx-68] r14
;   [rbx-20] r8    [rbx-48] rsi   [rbx-70] r15
;   [rbx-28] r9    [rbx-50] rdi   [rbx-78] flags
;
; Keep this table and the pushes in step: a wrong offset still reads valid memory, just the
; wrong register.

; The unwind info describes an empty prologue, so nothing may unwind through these frames;
; HookFilter in vbatrace.cpp handles every code. If that changes, add a .pushreg per push and
; .setframe rbx, 0 after `mov rbx, rsp`, before .endprolog, in the same change.
; A stack overflow in the saves is not caught; the fixed 264-byte frame keeps it within the
; interpreter's own margin.

option casemap:none

EXTERN XRayVbaOnStatement : PROC
EXTERN XRayVbaOnBreakpoint : PROC
EXTERN XRayVbaOnExit      : PROC
EXTERN XRayVbaOnEnd       : PROC
EXTERN XRayVbaOnStop      : PROC

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
; Entered by JMP, so the volatile-register rule does not apply: the hook clobbers xmm0-xmm5 and a
; live Double would come back wrong. MXCSR too, as CRT float formatting can change it.
; Layout above rsp, 16-aligned here so movaps is safe:
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
; Entered by CALL from a per-slot stub.
HOOK_THUNK MACRO thunkName, recorder
thunkName PROC FRAME
    .endprolog
    SAVE_ALL
    lea     rcx, [rbx + 16]     ; rcx = the interpreter's rsp at the dispatch
    mov     rdx, rbx            ; rdx = top of the saved register block
    and     rsp, -16
    SAVE_FP

    call    recorder
    RESTORE_FP
    RESTORE_ALL
    ret
thunkName ENDP
ENDM

HOOK_THUNK XRayVbaBosThunk,   XRayVbaOnStatement
; A statement with a breakpoint set: it stops in the editor after this, never reaching BoS's handler.
HOOK_THUNK XRayVbaBosBpThunk, XRayVbaOnBreakpoint
HOOK_THUNK XRayVbaExitThunk,  XRayVbaOnExit
; The End opcode (slot 619) fires no exit for the frames it kills, and runs before its handler.
HOOK_THUNK XRayVbaEndThunk,   XRayVbaOnEnd
; The Stop opcode (slot 613): the editor pauses here and resumes; no frame ends.
HOOK_THUNK XRayVbaStopThunk,  XRayVbaOnStop

END
