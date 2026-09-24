# Representative cases still pass with every case's VBA compiled into one module, since module
# content can matter to name resolution; each case file alone tests a module holding only itself.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# the case files are the single source of truth; each returns early under $StretchCollectOnly
$allCases = @()
$StretchCollectOnly = $true
foreach ($caseFile in (Get-ChildItem (Join-Path $PSScriptRoot 'cases') -Filter '*.test.ps1' | Sort-Object Name)) {
    $case = $null
    . $caseFile.FullName
    if ($case) { $allCases += $case }
}
$StretchCollectOnly = $false
$repNames = @('nesting-3', 'names-of-many', 'frame-types')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # cases assert names against the fixed leaf 'VbaRun.xlsm'
    New-XRayMacroBook $sx 'vbacc' @(@{ Kind = 1; Name = 'Cases'; Code = @($allCases | ForEach-Object { $_.Setup }) }) `
        -SheetName 'Sheet1' -Leaf 'VbaRun.xlsm'
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $bookPath = $book.Path; $bookLeaf = $book.Leaf

    $failCount = 0
    foreach ($repName in $repNames) {
        $c = @($allCases | Where-Object { $_.Name -eq $repName })[0]
        $why = $null

        $mark = Get-LogLength $paths.Log
        $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
        if ($pressed -ne 'pressed') { $why = "arm: $pressed" }
        if (-not $why) {
            $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
            if ($armLine -notmatch 'ARMED') { $why = "did not arm: $armLine" }
        }

        if (-not $why) {
            $ran = $true; $err = ''
            try {
                if ($c.Invoke.Formula) {
                    $ws.Range('H1').Formula = $c.Invoke.Formula
                    Invoke-XRayRecalc $app
                }
                else {
                    $qualified = "$bookLeaf!" + $c.Invoke.Name
                    if ($c.Invoke.Args.Count -eq 0) { $app.Run($qualified) | Out-Null }
                    else { $app.Run($qualified, $c.Invoke.Args[0]) | Out-Null }
                }
            } catch { $ran = $false; $err = $_.Exception.Message }

            $mark2 = Get-LogLength $paths.Log
            $lossy = Stop-XRayTrace $sx
            if ($lossy) { Complete-Test -Fail -Detail $lossy }
            $totLine = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2
            $names = Read-XRayNames $paths.Log $mark2
            $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)

            if (-not $ran -and -not $c.Invoke.MayRaise) { $why = "VBA call failed: $err" }
            elseif (-not $totLine) { $why = 'no totals line after disarm' }
            else {
                $t = ConvertFrom-XRayTotals $totLine
                if ($t.hookFaults -gt 0) { $why = "$($t.hookFaults) hook fault(s)" }
                elseif ($totLine -match 'BREAKER=OPEN') { $why = 'circuit breaker opened' }
                else {
                    $t | Add-Member -NotePropertyName names -NotePropertyValue $names -Force
                    $t | Add-Member -NotePropertyName rows -NotePropertyValue $rows -Force
                    $why = & $c.Expect $t
                }
            }
        }

        if ($why) { $failCount++; Write-TestCase -Name $repName -Fail -Detail $why }
        else { Write-TestCase -Name $repName -Pass }
    }

    if ($failCount -eq 0) {
        Complete-Test -Pass -Detail ("{0} representative case(s) held with all {1} setups co-compiled" -f $repNames.Count, $allCases.Count)
    }
    else {
        Complete-Test -Fail -Detail "$failCount representative case(s) failed under co-compilation"
    }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
