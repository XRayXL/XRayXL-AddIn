# In a project not compiled, a procedure is compiled the first time it runs, and until then its
# caller's pool entry names a stand-in with no types. So a parameter passed to it is untyped on the
# first call and typed on the next. The exit re-read names the slots as the entry did, so the
# first call's exit does not report the ByRef argument as changed.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

$moduleCode = @'
Public gN As Double

Public Function Helper(y As Double) As Double
    Helper = y * 2
End Function
Public Sub ExprPass(x As Double)
    gN = Helper(x)
End Sub

Public Sub Drive()
    Dim d As Double
    d = 2.5
    ExprPass d + 0
    ExprPass d + 0
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    New-XRayMacroBook $sx 'OnDemand' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { $app.Run($leaf + '!Drive') | Out-Null }
    $entries = @($s.Rows | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'ExprPass' })
    $exits   = @($s.Rows | Where-Object { $_.kind -eq 'exit'  -and $_.function -eq 'ExprPass' })
    foreach ($r in $entries) { Write-Output ("  entry [{0}] {1}" -f $r.typetext, (Remove-ArgAddress $r.args)) }
    foreach ($r in $exits)   { Write-Output ("  exit  [{0}]" -f $r.args) }

    Check 'both-calls-ran' ($entries.Count -eq 2 -and $exits.Count -eq 2) "entries $($entries.Count), exits $($exits.Count)"
    if ($entries.Count -eq 2) {
        Check 'the-first-call-is-untyped' (([string]$entries[0].typetext).StartsWith('?')) "first [$($entries[0].typetext)]"
        Check 'the-second-call-is-typed-by-its-now-compiled-callee' `
              (([string]$entries[1].typetext -replace '#.*$', '') -ceq 'Double&' -and (Remove-ArgAddress $entries[1].args) -ceq 'a1:Double&=2.5') `
              "second [$($entries[1].typetext)] $($entries[1].args)"
    }
    Check 'no-exit-reports-a-change-that-did-not-happen' (@($exits | Where-Object { $_.args }).Count -eq 0) `
          ("exit args: " + (($exits | ForEach-Object { "[$($_.args)]" }) -join ' '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("first [{0}] second [{1}]" -f $entries[0].typetext, $entries[1].typetext)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
