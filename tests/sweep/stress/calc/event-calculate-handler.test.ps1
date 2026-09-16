$case = @{ Name='event-calculate-handler'
     Modules=@{
       'M'=@'
Public cVol As Long
Public cCalc As Long
Public Function Vol() As Double
    Application.Volatile
    cVol = cVol + 1
    Vol = Rnd()
End Function
Public Function XR_Counts() As String
    XR_Counts = "Vol=" & cVol & ";Worksheet_Calculate=" & cCalc
End Function
'@
     }
     SheetCode=@'
Private Sub Worksheet_Calculate()
    Dim z As Long
    z = 1
    M.cCalc = M.cCalc + 1
End Sub
'@
     Cells=@{ 'A1'='=Vol()' }
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Calculate'; if ($why) { return $why }
        # How often a volatile cell and its Calculate event run is Excel's to decide; VBA counts both.
        Test-TracedCallsMatchCounters $t }
     Why='Worksheet_Calculate, which fires from inside the calc engine itself' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
