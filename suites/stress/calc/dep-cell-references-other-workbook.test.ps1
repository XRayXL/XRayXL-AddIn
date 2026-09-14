$case = @{ Name='dep-cell-references-other-workbook'
     Cells=@{ 'A1'='=''[DepG.xlsm]AUX''!A1 * 2' }
     Deps=@(
       @{ Name='DepG'; Modules=@{
         'MG'=@'
Public Function SrcUdf(ByVal n As Double) As Double
    Application.Volatile
    SrcUdf = n * 11
End Function
'@
       }; Cells=@{ 'A1'='=SrcUdf(3)' } }
     )
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'SrcUdf'; if ($why) { return $why }
        $null }
     Why='a formula depending on another books cell, which itself is a UDF result' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
