$case = @{ Name='function-with-return'
     Setup=@'
Public Function T_Add(ByVal a As Double, ByVal b As Double) As Double
    T_Add = a + b
End Function
Public Sub T_CallAdd()
    Dim d As Double
    d = T_Add(2, 3)
End Sub
'@
     Invoke=@{ Name='T_CallAdd'; Args=@() }
     Expect={ param($t)
        if ($t.procedures -lt 2) { return "expected 2 procedures, got $($t.procedures)" }
        if ($t.exits -lt 1)      { return "a Function returning a value should hit an exit slot" }
        $null }
     Why='a typed Function exit reaches one of the typed ExitProc slots' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
