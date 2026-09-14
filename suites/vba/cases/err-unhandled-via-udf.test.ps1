$case = @{ Name='err-unhandled-via-udf'
     Setup=@'
Public Function T_UHF() As Double
    T_UHF = T_UHF2()
End Function
Private Function T_UHF2() As Double
    T_UHF2 = T_UHF3()
End Function
Private Function T_UHF3() As Double
    Err.Raise 5, "T_UHF3", "deliberate"
End Function
'@
     Invoke=@{ Formula='=T_UHF()' }
     Expect={ param($t)
        # A TRULY unhandled unwind, driven from a cell rather than
        # Application.Run. Run puts up the modal VBA error dialog and waits for
        # a human, which is untestable from a script; a UDF that raises becomes
        # #VALUE! and unwinds silently. Same unwind, no dialog.
        #
        # Three frames are released at once with nothing executing afterwards,
        # so only the flush at Disarm can close them.
        if ($t.faults -gt 0) { return "$($t.faults) faults during an unhandled unwind" }
        if ($t.framesOpened -lt 1) { return "the UDF was never traced at all" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     Why='an error with NO handler anywhere, unwinding out of VBA entirely.
          Driven from a cell because Application.Run would block on a dialog' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
