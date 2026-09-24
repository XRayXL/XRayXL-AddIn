# The demo add-in's ReverseText reverses characters, not UTF-16 units: reversed unit by unit, an
# emoji's surrogate pair shows as two replacement characters. The trace must show the pair in order.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$xll = Join-Path $PSScriptRoot '..\..\..\build\x64\Release\DemoBehaviors\DemoBehaviors64.xll'
$smile = [char]::ConvertFromUtf32(0x1F600)
$text = 'ab ' + $smile + ' ' + [char]0x65E5
$want  = [char]0x65E5 + ' ' + $smile + ' ba'

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    if (-not $app.RegisterXLL((Resolve-Path $xll).Path)) { Complete-Test -Fail -Detail "RegisterXLL failed: $xll" }
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Value2 = $text
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $ws.Cells.Item(1, 2).Formula = '=ReverseText(A1)'
    Invoke-XRayRecalc $app
    $value = [string]$ws.Cells.Item(1, 2).Value2
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $exit = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId) |
              Where-Object { $_.kind -eq 'exit' -and $_.function -eq 'ReverseText' }) | Select-Object -Last 1

    Check 'the-cell-holds-the-characters-reversed' ($value -eq $want) `
          ("B1 code units: " + (($value.ToCharArray() | ForEach-Object { '{0:X4}' -f [int]$_ }) -join ' '))
    # Built from parts: the trace's escapes spelt out, so no editor turns them into characters.
    $bs = [string][char]0x5C
    $wantRet = '"' + $bs + 'u65E5 ' + $bs + 'uD83D' + $bs + 'uDE00 ba"'
    $gotRet = if ($exit) { ([string]$exit.ret).Trim() } else { '' }
    Check 'the-trace-shows-the-pair-in-order' ($gotRet -ceq $wantRet) `
          ("ret: [$gotRet] len $($gotRet.Length), want len $($wantRet.Length); codes: " +
           (($gotRet.ToCharArray() | ForEach-Object { '{0:X}' -f [int]$_ }) -join ' '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "ReverseText keeps a surrogate pair whole: $($exit.ret)"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
