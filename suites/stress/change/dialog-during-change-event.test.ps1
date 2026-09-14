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
     Why='the dialog appears while an EVENT frame is open: the nastiest teardown shape, because the handler is abandoned by End rather than returning' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
