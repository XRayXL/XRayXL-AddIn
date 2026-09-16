# XRayXL_IsArmed -- ASK BEFORE YOU SET, instead of setting and reading the echo.
#
# The setters refuse while anything is armed (the modes are read once, at
# arm, so a mid-session change would let the hot path see a value move under
# it). That refusal is an echoed STRING -- fine for a human, awkward for a
# caller, which had to attempt a set and parse the reply to find out it was not
# allowed. This answers directly.
#
# EITHER SOURCE COUNTS, because either one refuses a set. The function calls the
# same AnythingArmed() the setters test, so the two cannot disagree -- which is
# the property worth asserting, and this test asserts it by driving the setter
# and IsArmed against each other rather than by reading the source.
#
# VOLATILE MATTERS HERE. Arming happens through XRayXL_Arm, a different call, so
# a non-volatile cell would keep reporting the answer from whenever it last
# calculated. That is asserted from a CELL, not through Application.Run, since
# Run recalculates nothing and would pass either way.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    function IsArmed { return [bool]($app.Run('XRayXL_IsArmed')) }

    Check 'false-before-arming' (-not (IsArmed)) "IsArmed said '$(IsArmed)'"

    # A set must be ACCEPTED while it says false -- that is the contract the
    # caller is relying on.
    $echo = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'ALL'
    Check 'set-accepted-while-not-armed' ($echo -notmatch 'refused') $echo

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }

    Check 'true-once-armed' (IsArmed) "IsArmed said '$(IsArmed)' after $armLine"

    # ...and the refusal it predicts actually happens. Asserting the answer
    # without this would leave IsArmed free to be confidently wrong.
    $echo2 = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'TOP'
    Check 'set-refused-while-armed' ($echo2 -match 'refused') $echo2

    # ---- VOLATILE: a CELL must notice, without being edited ---------------
    $ws = $app.ActiveSheet
    $ws.Range('A1').Formula = '=XRayXL_IsArmed()'
    Invoke-XRayRecalc $app
    $whileArmed = $ws.Range('A1').Value2

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    Invoke-XRayRecalc $app
    $afterDisarm = $ws.Range('A1').Value2

    Check 'cell-said-true-while-armed'   (($whileArmed -is [bool]) -and $whileArmed)        "A1 held '$whileArmed'"
    Check 'cell-followed-the-disarm'     (($afterDisarm -is [bool]) -and -not $afterDisarm) "A1 held '$afterDisarm' after disarm"

    Check 'false-after-disarm' (-not (IsArmed)) "IsArmed said '$(IsArmed)'"
    $echo3 = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'ALL'
    Check 'set-accepted-again-after-disarm' ($echo3 -notmatch 'refused') $echo3

    # ---- the VBA source alone must also count -----------------------------
    # Either source refuses a set, so IsArmed has to be true when only VBA is
    # armed. A version that asked the XLL side only would pass everything above.
    # VBA can arm only once VBE7 is in the process, which nothing here forces.
    $vbeLoaded = $false
    try { $vbeLoaded = [bool]((Get-Process -Id $sx.ProcId).Modules | Where-Object { $_.ModuleName -ieq 'VBE7.DLL' }) } catch {}
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')
    $mark2 = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $vbaLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark2
    if ($vbaLine -match 'ARMED') {
        Check 'true-when-only-vba-is-armed' (IsArmed) "IsArmed='$(IsArmed)'  $vbaLine"
    } elseif ($vbeLoaded) {
        Check 'vba-arms-when-vbe7-is-loaded' $false "VBE7 is loaded, yet VBA did not arm: $vbaLine"
    } else {
        Write-XRayObservation 'vba-only-arm-not-exercised' "VBE7 is not loaded in this Excel, so VBA cannot arm: $vbaLine"
    }
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'IsArmed tracks both sources, is volatile in a cell, and agrees with the setters'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
