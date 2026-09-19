# A procedure with more parameters than the typetext column can spell.
#
# argcount is the count, not the length of the text: 60, never where the text ran out. The
# text stops after a whole name and says it was cut. Only a1 and a60 are used, so the rest
# read ?unseen: a type comes from the p-code that loads the parameter.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# Ten to a line: one line of sixty is past VBA's 1023-character limit.
$params = (0..5 | ForEach-Object { $r = $_; ((1..10) | ForEach-Object { "ByVal a$($r * 10 + $_) As Double" }) -join ', ' }) -join ", _`r`n    "
$values = (0..5 | ForEach-Object { $r = $_; ((1..10) | ForEach-Object { $r * 10 + $_ }) -join ', ' }) -join ", _`r`n    "
$moduleCode = @"
Public Sub Drive()
    Dim x As Double
    x = Sixty($values)
End Sub

Private Function Sixty($params) As Double
    Sixty = a1 + a60
End Function
"@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    New-XRayMacroBook $sx 'SixtyParams' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { $app.Run($leaf + '!Drive') | Out-Null }
    $e = EntryRowOf $s.Rows 'Sixty'

    Check 'argcount-is-sixty' ($e.argcount -eq '60') "argcount $($e.argcount)"
    Check 'the-typetext-says-it-was-cut' ($e.typetext -match '^Double(,[^,]+)*,\.\.\.$') "typetext $($e.typetext)"
    Check 'the-last-value-is-in-the-args' ($e.args -match '(^| )a60:\S+=60( |$)') "args end: $(if ($e.args.Length -gt 60) { $e.args.Substring($e.args.Length - 60) } else { $e.args })"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "argcount $($e.argcount); typetext $($e.typetext.Length) chars"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
