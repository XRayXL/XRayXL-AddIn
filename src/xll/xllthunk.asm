; XRayXL -- the shared wrapping thunk for a traced XLL function.
;
; Every hooked function enters here with r10 pointing at its own Target. We
; record the entry, call the ORIGINAL function, record the exit, and return
; whatever it returned.
;
; A real call, not a rewritten return address: overwriting the return address mismatches the CET
; shadow stack, which is a fast-fail that bypasses SEH. An ordinary `call` is matched on `ret`.
;
; Stack arguments: the count comes from the registration string, so exactly that many are
; copied. Reading past the caller's frame can walk off the end of a stack near its limit.
;
; Registers: only volatile registers are scratch. rbx and rbp are taken and pushed; rsi/rdi
; would corrupt the caller.
;
; FRAME. A frame pointer, with the outgoing area sized per call to the signature
; (up to Excel's 255 arguments). .setframe keeps it unwindable.
;
;   entry rsp = 8 (mod 16)          [return address pushed by the caller]
;   push rbp                     -> 0 (mod 16)
;   push rbx                     -> 8 (mod 16)
;   sub rsp, 68h                 -> 0 (mod 16)
;   mov rbp, rsp                    the frame base, 16-aligned
;
;   rbp+000h   Regs (60h)
;   rbp+060h   the clamped stack-argument count, kept across our calls
;   rbp+068h   where rsp was after the two pushes -- the epilogue's target
;   rbp+0A0h   the caller's fifth argument (see below)
;
;   then, dynamically:
;   rsp+000h   32 bytes shadow space for the calls we make
;   rsp+020h   outgoing stack arguments, as many as this signature has
;
; Where the caller's fifth argument is:
;   at our entry            rsp_e -> [rsp_e+00h] return address
;                                    [rsp_e+08h..27h] 32 bytes of SHADOW SPACE
;                                    [rsp_e+28h] argument 4   <-- the fifth
;   push rbp, push rbx      rsp = rsp_e - 10h
;   sub rsp, 68h            rsp = rsp_e - 78h = rbp
;   so argument 4 sits at   rbp + 78h + 28h = rbp + 0A0h
;
; Outgoing, the symmetry is different and must not be copied from the above:
; our `call` pushes 8 more, so the callee reads its argument 4 at ITS rsp+28h,
; which is our rsp+20h -- FRAME_OUTARGS.
;
; STACK PROBING. The outgoing area can exceed a page, so it is probed page by page.
;
; Regs layout is shared with xllregs.h, which static_asserts these offsets.

OPTION CASEMAP:NONE

; --- Target fields this file touches (see xllhook.h) ------------------------
TGT_ORIGINAL    EQU 0           ; void*  the real function
TGT_ARGCOUNT    EQU 8           ; int32  ABI argument slots

; --- Regs offsets (xllregs.h) -----------------------------------------------
R_IREG          EQU 000h
R_XMM           EQU 020h
R_RAX           EQU 040h
R_XMMRET        EQU 048h
R_STACKARGS     EQU 050h

; --- our frame, all relative to rbp -----------------------------------------
FIXED_SIZE      EQU 068h        ; Regs (60h) + the saved count (8h)
F_NSTACK        EQU 060h        ; the clamped stack-argument count
FRAME_OUTARGS   EQU 020h        ; from rsp, after the dynamic allocation
CALLER_ARG4     EQU 0A0h        ; from rbp: 78h to entry rsp, plus 28h
SHADOW_SPACE    EQU 020h
PAGE_SIZE       EQU 1000h
; Plan::kMaxAbiSlots (768) less the four register arguments.
MAX_STACK_ARGS  EQU 764

EXTERN XRayOnEntry:PROC           ; void __fastcall XRayOnEntry(Target*, Regs*)
EXTERN XRayOnExit:PROC            ; void __fastcall XRayOnExit (Target*, Regs*)

.CODE

XRayXllThunk PROC FRAME
        push    rbp
        .pushreg rbp
        push    rbx
        .pushreg rbx
        sub     rsp, FIXED_SIZE
        .allocstack FIXED_SIZE
        mov     rbp, rsp
        .setframe rbp, 0
        .endprolog

        mov     rbx, r10                ; Target*, across both our calls

        ; ---- snapshot the incoming arguments -------------------------------
        mov     [rbp+R_IREG+00h], rcx
        mov     [rbp+R_IREG+08h], rdx
        mov     [rbp+R_IREG+10h], r8
        mov     [rbp+R_IREG+18h], r9
        movq    QWORD PTR [rbp+R_XMM+00h], xmm0
        movq    QWORD PTR [rbp+R_XMM+08h], xmm1
        movq    QWORD PTR [rbp+R_XMM+10h], xmm2
        movq    QWORD PTR [rbp+R_XMM+18h], xmm3

        lea     rax, [rbp+CALLER_ARG4]
        mov     [rbp+R_STACKARGS], rax

        ; ---- how many stack arguments this signature has --------------------
        ; Clamped: a stale Target must not cause an unbounded allocation.
        mov     eax, DWORD PTR [rbx+TGT_ARGCOUNT]
        sub     eax, 4
        jg      have_stack_args
        xor     eax, eax
have_stack_args:
        cmp     eax, MAX_STACK_ARGS
        jbe     count_ok
        mov     eax, MAX_STACK_ARGS
count_ok:
        mov     DWORD PTR [rbp+F_NSTACK], eax

        ; ---- allocate shadow space and the outgoing arguments ---------------
        mov     ecx, eax                ; zero-extends
        shl     rcx, 3                  ; slots -> bytes
        add     rcx, 0Fh
        and     rcx, -16                ; 16-aligned, so rsp stays aligned
        add     rcx, SHADOW_SPACE

        ; Probe every page, including the last partial one.
        mov     r10, rcx                ; bytes still to probe
        mov     r11, rsp
probe_loop:
        cmp     r10, PAGE_SIZE
        jb      probe_last
        sub     r11, PAGE_SIZE
        mov     BYTE PTR [r11], 0
        sub     r10, PAGE_SIZE
        jmp     probe_loop
probe_last:
        test    r10, r10
        jz      probe_done
        sub     r11, r10
        mov     BYTE PTR [r11], 0
probe_done:
        sub     rsp, rcx

        ; ---- record the entry ----------------------------------------------
        mov     rcx, rbx
        mov     rdx, rbp
        call    XRayOnEntry

        ; ---- copy the stack arguments we were told to expect -----------------
        movsxd  rcx, DWORD PTR [rbp+F_NSTACK]
        test    rcx, rcx
        jz      no_stack_args
        lea     r10, [rbp+CALLER_ARG4]          ; source
        lea     r11, [rsp+FRAME_OUTARGS]        ; destination
copy_loop:
        mov     rax, [r10]
        mov     [r11], rax
        add     r10, 8
        add     r11, 8
        dec     rcx
        jnz     copy_loop
no_stack_args:

        ; ---- restore the registers exactly as they arrived -------------------
        mov     rcx, [rbp+R_IREG+00h]
        mov     rdx, [rbp+R_IREG+08h]
        mov     r8,  [rbp+R_IREG+10h]
        mov     r9,  [rbp+R_IREG+18h]
        movq    xmm0, QWORD PTR [rbp+R_XMM+00h]
        movq    xmm1, QWORD PTR [rbp+R_XMM+08h]
        movq    xmm2, QWORD PTR [rbp+R_XMM+10h]
        movq    xmm3, QWORD PTR [rbp+R_XMM+18h]

        ; ---- the real call. An ordinary call/ret pair. -----------------------
        call    QWORD PTR [rbx+TGT_ORIGINAL]

        ; ---- snapshot the result --------------------------------------------
        mov     [rbp+R_RAX], rax
        movq    QWORD PTR [rbp+R_XMMRET], xmm0

        mov     rcx, rbx
        mov     rdx, rbp
        call    XRayOnExit

        ; ---- hand back exactly what the callee returned ----------------------
        mov     rax, [rbp+R_RAX]
        movq    xmm0, QWORD PTR [rbp+R_XMMRET]

        lea     rsp, [rbp+FIXED_SIZE]
        pop     rbx
        pop     rbp
        ret
XRayXllThunk ENDP

END
