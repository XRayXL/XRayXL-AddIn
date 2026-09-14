$case = @{ Name='event-change-calculate-with-error'
     Modules=@{
       'M'=@'
Public Function BadUdf(ByVal n As Double) As Double
    If n > 1 Then Err.Raise 5401
    BadUdf = n
End Function
'@
     }
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    Me.UsedRange.Dirty
    Application.Calculate
    Application.EnableEvents = True
    Err.Clear
End Sub
'@
     Cells=@{ 'A1'='=BadUdf(1)'; 'A2'='=BadUdf(2)' }
     Trigger=@{ Kind='Change'; Cell='B2'; Value=5 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','BadUdf'; if ($why) { return $why }
        $null }
     Why='the Change handler recalculates a sheet whose UDF raises: an error crossing the re-entry' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
