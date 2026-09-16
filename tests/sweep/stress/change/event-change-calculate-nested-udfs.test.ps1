$case = @{ Name='event-change-calculate-nested-udfs'
     Modules=@{
       'M'=@'
Public cA As Long
Public cB As Long
Public cC As Long
Public cChange As Long
Public Function UdfA(ByVal n As Double) As Double
    cA = cA + 1
    UdfA = UdfB(n) + 1
End Function
Public Function UdfB(ByVal n As Double) As Double
    cB = cB + 1
    UdfB = UdfC(n) + 1
End Function
Public Function UdfC(ByVal n As Double) As Double
    cC = cC + 1
    UdfC = n * 2
End Function
Public Function XR_Counts() As String
    XR_Counts = "UdfA=" & cA & ";UdfB=" & cB & ";UdfC=" & cC & ";Worksheet_Change=" & cChange
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
     Cells=@{ 'A1'='=UdfA(10)'; 'A2'='=UdfA(20)' }
     Trigger=@{ Kind='Change'; Cell='B2'; Value=3 }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','UdfA','UdfB','UdfC'; if ($why) { return $why }
        if ($t.maxDepth -lt 3) { return "expected nesting >=3, got $($t.maxDepth)" }
        # How many times the handler's Calculate evaluates the sheet is Excel's to decide; VBA counts it.
        Test-TracedCallsMatchCounters $t }
     Why='a Change handler recalculating a sheet whose UDFs call each other three deep' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
