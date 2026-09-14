# IN PROCESS: arming must report the XLL functions it armed, and the VBA
# dispatch-table derivation must agree with reality about whether VBE7 is even
# loaded. This is the only place the derivation runs against a VBE7 that Excel
# itself loaded, at the address Excel put it, rather than against a file.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $allGood = $true

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }

    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    $ok = [bool]($armLine -match 'armed \d+ of')
    if (-not $ok) { $allGood = $false }
    Write-TestCase -Name 'xll-functions-armed' -Pass:$ok -Fail:(-not $ok) -Detail "$armLine"

    $vbaLine = Wait-LogLine $paths.Log 'VBA:' $mark 30
    if (-not $vbaLine) {
        Write-TestCase -Name 'vba-derivation-ran' -Fail -Detail 'no VBA: line in the action log'
        Complete-Test -Fail -Detail 'the step-1 derivation did not run'
    }
    Write-TestCase -Name 'vba-derivation-ran' -Pass -Detail ($vbaLine.Trim())

    # Deterministic without depending on VBA trust settings: ask the PROCESS
    # whether VBE7 is loaded, then require the log to agree with the answer.
    $vbeLoaded = $false
    try {
        $vbeLoaded = [bool]((Get-Process -Id $sx.ProcId).Modules |
                            Where-Object { $_.ModuleName -ieq 'VBE7.DLL' })
    } catch {}

    if ($vbeLoaded) {
        # VBA is in the process: a refusal is a real regression, not a shrug.
        # the table size is structural; how many distinct handlers fill it differs between VBE7 builds
        $ok2 = ($vbaLine -match 'VERIFIED') -and ($vbaLine -match 'slots=1700\b') -and ($vbaLine -match 'distinct=[1-9]\d*')
        if (-not $ok2) { $allGood = $false }
        Write-TestCase -Name 'vbe7-loaded-table-verified' -Pass:$ok2 -Fail:(-not $ok2) -Detail ($vbaLine.Trim())
    }
    else {
        # VBA is NOT loaded: the only correct answer is to say so -- a table
        # reported here would mean the derivation invented one.
        $ok2 = [bool]($vbaLine -match 'not loaded')
        if (-not $ok2) { $allGood = $false }
        Write-TestCase -Name 'vbe7-absent-correctly-declined' -Pass:$ok2 -Fail:(-not $ok2) -Detail ($vbaLine.Trim())
    }

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    if ($allGood) { Complete-Test -Pass -Detail "vbe7Loaded=$vbeLoaded" }
    else { Complete-Test -Fail -Detail 'one or more derivation checks failed' }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
