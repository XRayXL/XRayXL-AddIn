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
        # From a cell because under Application.Run the same unwind waits on a modal dialog.
        # Nothing runs afterwards, so only the flush at Disarm can close the three frames.
        if ($t.faults -gt 0) { return "$($t.faults) faults during an unhandled unwind" }
        if ($t.framesOpened -lt 1) { return "the UDF was never traced at all" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     # Entering the formula evaluates H1 once and the driver's CalculateFull again: two chains.
     # The function Excel entered for the cell reads unhandled; the frames below keep threw and unwound.
     Calls=@(
        @{ Function='T_UHF';  Depth='1'; Parent=-1; Caller='cell'; Cell='H1'; Outcome='unhandled' }
        @{ Function='T_UHF2'; Depth='2'; Parent=0; Outcome='unwound' }
        @{ Function='T_UHF3'; Depth='3'; Parent=1; Outcome='threw' }
        @{ Function='T_UHF';  Depth='1'; Parent=-1; Caller='cell'; Cell='H1'; Outcome='unhandled' }
        @{ Function='T_UHF2'; Depth='2'; Parent=3; Outcome='unwound' }
        @{ Function='T_UHF3'; Depth='3'; Parent=4; Outcome='threw' } )
     Why='an error with NO handler anywhere, unwinding out of VBA entirely.
          Driven from a cell because Application.Run would block on a dialog' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
