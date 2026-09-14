$case = @{ Name='event-change-writes-500-cells'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim i As Long
    Application.EnableEvents = False
    For i = 1 To 500
        Me.Cells(i, 8).Value = i
    Next i
    Application.EnableEvents = True
End Sub
'@
     Trigger=@{ Kind='Change'; Cell='B2'; Value=9 }
     Expect={ param($t)
        if ($t.statements -lt 500) { return "expected >=500 statements, got $($t.statements)" }
        $null }
     Why='a Change handler writing 500 cells, each write a potential re-entry point' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
