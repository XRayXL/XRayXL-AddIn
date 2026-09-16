$case = @{ Name='result-currency-and-dates'
     Modules=@{
       'M'=@'
Public Sub Go()
    Dim ws As Object
    Dim c As Currency
    Dim d As Date
    Set ws = ThisWorkbook.Worksheets("S1")
    c = -0.5
    ws.Range("D1").Value = c
    c = 1234.5678
    ' MEASURED, with no tracing at all: assigning a Currency straight to a
    ' cell coerces to TWO decimals (1234.57). CDbl first and the full four
    ' survive. Both are recorded so the case tests Excel's real behaviour
    ' rather than my prediction of it.
    ws.Range("D2").Value = c
    ws.Range("D4").Value = CDbl(c)
    d = DateSerial(2020, 1, 2)
    ws.Range("D3").Value = CDbl(d)
    Call TakesMoney(-0.5, 1234.5678)
End Sub
Public Sub TakesMoney(ByVal a As Currency, ByVal b As Currency)
    Dim z As Currency
    z = a
    z = b
End Sub
'@
     }
     VerifyCells=@{ 'D1'='-0.5'; 'D2'='1234.57'; 'D4'='1234.5678'; 'D3'='43832' }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.statements -lt 5) { return "the VBA did not run" }
        $null }
     # Currency is written with its four decimals, the sign kept.
     Calls=@(
        @{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }
        @{ Function='TakesMoney'; Args='a1:Currency=-0.5000 a2:Currency=1234.5678'; Depth='2'; Parent=0; Outcome='returned' } )
     Why='Currency and Date values, including NEGATIVE currency, must survive tracing' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
