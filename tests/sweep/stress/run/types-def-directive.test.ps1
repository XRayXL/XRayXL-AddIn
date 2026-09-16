# `DefLng A-C` types an untyped parameter by first letter, far from the declaration
# (MS-VBAL 5.2.3.1.5). It should decode as `As Long`. Own module: Def directives are module-scoped.
$case = @{ Name='types-def-directive'
     Modules=@{
       'M'=@'
DefLng Q
DefStr W
DefDbl E
Public Sub DefLongParam(ByVal qVal)
    Dim z As Long
    z = qVal
End Sub
Public Sub DefStrParam(ByVal wVal)
    Dim z As String
    z = wVal
End Sub
Public Sub DefDblParam(ByVal eVal)
    Dim z As Double
    z = eVal
End Sub
Public Sub NoDirectiveParam(ByVal pVal)
    Dim z As Variant
    z = pVal
End Sub
Public Sub Go()
    Call DefLongParam(&H11223344)
    Call DefStrParam("DEFSTR")
    Call DefDblParam(2748.5)
    Call NoDirectiveParam(42)
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = Get-FirstEntryByName $t.rows
        $want = @{ 'DefLongParam'='Long'; 'DefStrParam'='String';
                   'DefDblParam'='Double'; 'NoDirectiveParam'='Variant' }
        foreach ($fn in $want.Keys) {
            if (-not $e.ContainsKey($fn)) { return "$fn was not traced" }
            if ($e[$fn].typetext -ne $want[$fn]) {
                return "$fn signature was [$($e[$fn].typetext)], expected $($want[$fn]) -- a Def directive changes the declared type of an untyped parameter" }
        }
        $null }
     Why='Def directives set an untyped parameters type from a module-level
          declaration far from the parameter itself -- the only form where the
          declaration is not visible at the parameter' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
