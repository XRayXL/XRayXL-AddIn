$case = @{ Name='dep-cell-references-other-workbook'
     Modules=@{
       'M'=@'
Public Function XR_Counts() As String
    ' The counted UDF lives in the other book; ask it, so one macro answers for both.
    XR_Counts = Application.Run("'DepG.xlsm'!XR_SrcCounts")
End Function
'@
     }
     Cells=@{ 'A1'='=''[DepG.xlsm]AUX''!A1 * 2' }
     Deps=@(
       @{ Name='DepG'; Modules=@{
         'MG'=@'
Public cSrc As Long
Public Function SrcUdf(ByVal n As Double) As Double
    Application.Volatile
    cSrc = cSrc + 1
    SrcUdf = n * 11
End Function
Public Function XR_SrcCounts() As String
    XR_SrcCounts = "SrcUdf=" & cSrc
End Function
'@
       }; Cells=@{ 'A1'='=SrcUdf(3)' } }
     )
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'SrcUdf'; if ($why) { return $why }
        # A volatile UDF in another book runs as often as Excel decides; its own book counts it.
        Test-TracedCallsMatchCounters $t }
     Why='a formula depending on another books cell, which itself is a UDF result' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
