; A second shape of export. Many XLLs export a jump table: the exported symbol is a six-byte
; `jmp qword ptr [rip+disp32]` and the real address lives in a writable data slot. Incremental
; linking and import thunks produce this, and the tracer must handle both shapes. Per export:
;
;    .code    FF 25 <disp32>        jmp qword ptr [rip+disp32]     6 bytes
;    .data    <qword>               -> the real implementation      writable
;
; The implementations are the ordinary C++ functions in tracedaddin.cpp.

OPTION CASEMAP:NONE

EXTERN TxJumpAddImpl:PROC
EXTERN TxJumpStrImpl:PROC

.DATA
; The slot table. Writable by definition -- it is .data -- which is what makes
; redirecting one an aligned pointer store rather than a code patch.
PUBLIC TxJumpSlots
TxJumpSlots     LABEL QWORD
SlotAdd         QWORD TxJumpAddImpl
SlotStr         QWORD TxJumpStrImpl

.CODE

; Each of these assembles to exactly FF 25 disp32.
PUBLIC TxJumpAdd
TxJumpAdd PROC
        jmp     QWORD PTR [SlotAdd]
TxJumpAdd ENDP

PUBLIC TxJumpStr
TxJumpStr PROC
        jmp     QWORD PTR [SlotStr]
TxJumpStr ENDP

END
