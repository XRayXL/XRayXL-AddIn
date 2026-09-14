$case = @{ Name='event-change-calculate-nested-udfs'
     Modules=@{
       'M'=@'
Public Function UdfA(ByVal n As Double) As Double
    UdfA = UdfB(n) + 1
End Function
Public Function UdfB(ByVal n As Double) As Double
    UdfB = UdfC(n) + 1
End Function
Public Function UdfC(ByVal n As Double) As Double
    UdfC = n * 2
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
     Cells=@{ 'A1'='=UdfA(10)'; 'A2'='=UdfA(20)' }
     Trigger=@{ Kind='Change'; Cell='B2'; Value=3 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','UdfA','UdfB','UdfC'; if ($why) { return $why }
        if ($t.maxDepth -lt 3) { return "expected nesting >=3, got $($t.maxDepth)" }
        $null }
     Why='a Change handler recalculating a sheet whose UDFs call each other three deep' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
