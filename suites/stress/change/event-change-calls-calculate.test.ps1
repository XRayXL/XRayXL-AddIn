$case = @{ Name='event-change-calls-calculate'
     Modules=@{
       'M'=@'
Public Function SlowUdf(ByVal n As Double) As Double
    Dim i As Long
    Dim s As Double
    For i = 1 To 50
        s = s + i
    Next i
    SlowUdf = s + n
End Function
'@
     }
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Me.UsedRange.Dirty
    Application.Calculate
    Application.EnableEvents = True
End Sub
'@
     Cells=@{ 'A1'='=SlowUdf(1)'; 'A2'='=SlowUdf(2)'; 'A3'='=SlowUdf(3)' }
     Trigger=@{ Kind='Change'; Cell='B2'; Value=7 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','SlowUdf'; if ($why) { return $why }
        if ($t.maxDepth -lt 2) { return "re-entry produced no nesting (depth $($t.maxDepth))" }
        $null }
     Why='THE re-entrancy case: a Change handler calls Calculate, so the interpreter re-enters through a different trigger while the handler frame is still open' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
