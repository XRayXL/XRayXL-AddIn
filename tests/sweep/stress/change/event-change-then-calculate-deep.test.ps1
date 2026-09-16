$case = @{ Name='event-change-then-calculate-deep'
     Modules=@{
       'M'=@'
Public cRec As Long
Public cChange As Long
Public Function RecUdf(ByVal n As Double) As Double
    cRec = cRec + 1
    If n <= 0 Then
        RecUdf = 0
    Else
        RecUdf = n + RecUdf(n - 1)
    End If
End Function
Public Function XR_Counts() As String
    XR_Counts = "RecUdf=" & cRec & ";Worksheet_Change=" & cChange
End Function
'@
     }
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    M.cChange = M.cChange + 1
    Application.EnableEvents = False
    Me.UsedRange.Dirty
    Application.Calculate
    Application.EnableEvents = True
End Sub
'@
     Cells=@{ 'A1'='=RecUdf(8)' }
     Trigger=@{ Kind='Change'; Cell='B5'; Value=4 }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','RecUdf'; if ($why) { return $why }
        if ($t.recursions -lt 5) { return "expected recursion inside the re-entry, got $($t.recursions)" }
        # Every activation counts itself, recursion included, so the rows must match exactly.
        Test-TracedCallsMatchCounters $t }
     Why='Change -> Calculate -> UDF -> recursive UDF: re-entry AND recursion together' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
