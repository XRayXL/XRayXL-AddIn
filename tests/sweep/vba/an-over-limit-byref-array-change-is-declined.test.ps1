# A ByRef array too big to write out, changed by the callee, counts as declined, not unchanged:
# there is no entry value to compare. The refusal log counts only arrays a reader lost.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub Drive()
    Dim big() As Double
    ReDim big(0 To 3000000)
    TakeBig big
    Dim v As Variant
    v = BigBack()
End Sub

Private Sub TakeBig(ByRef a() As Double)
    a(5) = 42
End Sub

Private Function BigBack() As Variant
    Dim a() As Double
    ReDim a(0 To 3000000)
    BigBack = a
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $log = (Get-XRayPaths $sx.ProcId).Log

    New-XRayMacroBook $sx 'OverLimitByRef' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $log
    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { $app.Run($leaf + '!Drive') | Out-Null }
    $valuesLine = @(Get-Content $log | Select-Object -Skip $mark | Where-Object { $_ -match 'values: \d+ array' }) | Select-Object -Last 1
    $t = ConvertFrom-XRayTotals $s.Totals

    $exit = ExitRowOf $s.Rows 'TakeBig'
    Write-XRayObservation 'entry-args' (ArgsOf $s.Rows 'TakeBig')
    Check 'the-change-is-declined-not-same' (($t.byrefDeclined -ge 1) -and ($t.byrefSame -eq 0)) "$($s.Totals)"
    Check 'no-exit-args-claim-a-value' ([string]::IsNullOrEmpty($exit.args)) "exit args: '$($exit.args)'"
    Check 'the-refusal-counts-the-entry-and-the-return' ($valuesLine -match 'values: 2 array') "$valuesLine"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "byrefDeclined=$($t.byrefDeclined); $valuesLine"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
