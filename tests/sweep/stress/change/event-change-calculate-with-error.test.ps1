$case = @{ Name='event-change-calculate-with-error'
     Modules=@{
       'M'=@'
Public cBad As Long
Public cChange As Long
Public Function BadUdf(ByVal n As Double) As Double
    cBad = cBad + 1
    If n > 1 Then Err.Raise 5401
    BadUdf = n
End Function
Public Function XR_Counts() As String
    XR_Counts = "BadUdf=" & cBad & ";Worksheet_Change=" & cChange
End Function
'@
     }
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    M.cChange = M.cChange + 1
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
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','BadUdf'; if ($why) { return $why }
        # The counter is incremented BEFORE the raise, so a raising call is counted too:
        # every call must have its row, whether it returned or threw.
        Test-TracedCallsMatchCounters $t }
     Why='the Change handler recalculates a sheet whose UDF raises: an error crossing the re-entry' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
