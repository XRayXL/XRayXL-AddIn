# A PROCEDURE NAMED OUTSIDE ASCII READS AS ITSELF.
#
# VBA keeps identifiers in the ANSI code page and the trace file is UTF-8, so a
# name copied across byte for byte is not valid UTF-8 and reads as replacement
# characters. The summary grid Excel displays must agree with the file.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# Built from code points, so this file's own encoding cannot change the name.
$name = 'Gr' + [char]0x00F6 + [char]0x00DF + 'e'

$moduleCode = @"
Public Function $name(ByVal a As Long) As Long
    $name = a + 1
End Function

Public Sub NA_Drive()
    Dim v As Long
    v = $name(1)
End Sub
"@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'NonAscii' @(
        @{ Kind=1; Name='NonAsciiCase'; Code=$moduleCode }
    )
    $leaf = (Get-XRayMacroBook).Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!NA_Drive') | Out-Null
    $summary = ConvertFrom-XRayTraceSummary ($app.Run('XRayXL_GetTraceSummary'))
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $csv = Get-XRayTraceCsv $sx.ProcId
    $rows = @(Read-TraceRows $sx.ProcId)
    $seen = @($rows | Where-Object { $_.source -eq 'VBA' -and $_.kind -eq 'entry' } |
              ForEach-Object { $_.function }) -join ','
    $listed = @($summary.Rows | ForEach-Object { $_.Function }) -join ','

    # The bytes first, so the verdict does not rest on how the reader decodes.
    $want = [Text.Encoding]::UTF8.GetBytes($name)
    $bytes = [IO.File]::ReadAllBytes($csv)
    $hex = { param($b) ($b | ForEach-Object { $_.ToString('X2') }) -join '' }
    $inFile = (& $hex $bytes).Contains((& $hex $want))

    Check 'the-file-holds-the-name-as-utf8' $inFile "UTF-8 bytes $(& $hex $want) not found in $csv"
    Check 'the-trace-row-names-it' (@($rows | Where-Object { $_.function -eq $name }).Count -ge 1) `
          "entry functions: $seen"
    Check 'the-summary-names-it' (@($summary.Rows | Where-Object { $_.Function -eq $name }).Count -eq 1) `
          "summary functions: $listed"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "named in the file and the summary"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
