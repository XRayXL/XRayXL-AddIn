$case = @{ Name='event-calculate-handler'
     Modules=@{
       'M'=@'
Public Function Vol() As Double
    Application.Volatile
    Vol = Rnd()
End Function
'@
     }
     SheetCode=@'
Private Sub Worksheet_Calculate()
    Dim z As Long
    z = 1
End Sub
'@
     Cells=@{ 'A1'='=Vol()' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Calculate'; if ($why) { return $why }
        $null }
     Why='Worksheet_Calculate, which fires from inside the calc engine itself' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
