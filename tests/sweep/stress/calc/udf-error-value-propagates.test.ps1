$case = @{ Name='udf-error-value-propagates'
     Modules=@{
       'M'=@'
Private cE As Long
Public Function ErrUdf(ByVal n As Double) As Variant
    cE = cE + 1
    If n > 3 Then
        ErrUdf = CVErr(2042)
    Else
        ErrUdf = n
    End If
End Function
Public Function XR_Counts() As String
    XR_Counts = "ErrUdf=" & cE
End Function
'@
     }
     FillFormula=@{ Range='A1:A10'; Formula='=ErrUdf(ROW())' }
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'ErrUdf'; if ($why) { return $why }
        # Returning CVErr raises nothing, so every call still reads returned.
        $notReturned = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.function -eq 'ErrUdf' -and $_.outcome -ne 'returned' })
        if ($notReturned.Count) {
            return ("a UDF returning an error value is not a raise: " +
                    (($notReturned | ForEach-Object { $_.outcome }) -join ',')) }
        Test-TracedCallsMatchCounters $t }
     Why='a UDF returning CVErr, and another consuming the error value' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
