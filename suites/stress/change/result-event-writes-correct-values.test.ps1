$case = @{ Name='result-event-writes-correct-values'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    Me.Range("D1").Value = 2# / 3#
    Me.Range("D2").Value = Sqr(2#)
    Application.EnableEvents = True
End Sub
'@
     VerifyCells=@{ 'D1'='0.666666666666667'; 'D2'='1.4142135623731' }
     Trigger=@{ Kind='Change'; Cell='B2'; Value=1 }
     Expect={ param($t)
        if ($t.statements -lt 5) { return "the handler did not run" }
        $null }
     Why='a Change handler computing Doubles must write the right ones while traced' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
