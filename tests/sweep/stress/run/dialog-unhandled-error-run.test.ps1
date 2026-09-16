$case = @{ Name='dialog-unhandled-error-run'
     Modules=@{
       'M'=@'
Public Sub RaiseLeaf()
    Err.Raise 5901, "RaiseLeaf", "unhandled on purpose"
End Sub
Public Sub RaiseMid()
    RaiseLeaf
End Sub
Public Sub RaiseTop()
    RaiseMid
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='RaiseTop'; MayRaise=$true; ExpectDialog=$true }
     Expect={ param($t)
        # MEASURED: an unhandled Err.Raise under Application.Run DOES raise
        # Excel's modal VBA dialog, which blocks the calling thread until the
        # watchdog presses End. Frames abandoned that way fire no exit opcode,
        # so the stack-pointer backstop is what has to balance them.
        if ($t.dialogs -lt 1) { return "expected the modal VBA dialog, none appeared" }
        if ($t.framesOpened -lt 3) { return "expected >=3 frames, got $($t.framesOpened)" }
        $why = Assert-VbaTraced $t 'RaiseTop','RaiseMid','RaiseLeaf'; if ($why) { return $why }
        $null }
     # One chain, three deep. NO outcome expected: the dialog's End button kills these frames,
     # and what a frame closed that way reads is a documented limitation (docs/TraceRowModel.md).
     Calls=@(
        @{ Function='RaiseTop';  Depth='1'; Parent=-1 }
        @{ Function='RaiseMid';  Depth='2'; Parent=0 }
        @{ Function='RaiseLeaf'; Depth='3'; Parent=1 } )
     Why='an unhandled error under Application.Run raises Excels modal VBA dialog; the watchdog presses End, which abandons frames without any exit opcode' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
