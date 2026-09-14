$case = @{ Name='event-change-cascade-guarded'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Static depth As Long
    If depth > 3 Then Exit Sub
    depth = depth + 1
    Me.Range("C" & depth).Value = depth
    depth = depth - 1
End Sub
'@
     Trigger=@{ Kind='Change'; Cell='B2'; Value=1 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change'; if ($why) { return $why }
        $null }
     Why='a Change handler that WRITES cells, re-entering itself until a guard stops it' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
