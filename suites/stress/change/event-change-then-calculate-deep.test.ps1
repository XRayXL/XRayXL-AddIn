$case = @{ Name='event-change-then-calculate-deep'
     Modules=@{
       'M'=@'
Public Function RecUdf(ByVal n As Double) As Double
    If n <= 0 Then
        RecUdf = 0
    Else
        RecUdf = n + RecUdf(n - 1)
    End If
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
     Cells=@{ 'A1'='=RecUdf(8)' }
     Trigger=@{ Kind='Change'; Cell='B5'; Value=4 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','RecUdf'; if ($why) { return $why }
        if ($t.recursions -lt 5) { return "expected recursion inside the re-entry, got $($t.recursions)" }
        $null }
     Why='Change -> Calculate -> UDF -> recursive UDF: re-entry AND recursion together' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
