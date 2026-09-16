$case = @{ Name='dialog-during-change-event'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim z As Long
    z = 1
    Boom
End Sub
Private Sub Boom()
    Err.Raise 5902
End Sub
'@
     Trigger=@{ Kind='Change'; Cell='B9'; Value=77; MayRaise=$true; ExpectDialog=$true }
     Expect={ param($t)
        if ($t.dialogs -lt 1) { return "no dialog recorded" }
        $why = Assert-VbaTraced $t 'Worksheet_Change'; if ($why) { return $why }
        $null }
     # The handler and the raise, nothing else. NO outcome expected: the dialog's End button
     # kills both frames, and what they then read is a documented limitation (docs/TraceRowModel.md).
     Calls=@(
        @{ Function='Worksheet_Change'; Depth='1'; Parent=-1 }
        @{ Function='Boom';             Depth='2'; Parent=0 } )
     Why='the dialog appears while an EVENT frame is open: the nastiest teardown shape, because the handler is abandoned by End rather than returning' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
