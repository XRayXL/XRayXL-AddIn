# The setter surface is observable: every set echoes what it did, XRayXL_GetTraceParam answers
# the current value, and an off-vocabulary value is refused with nothing changed. A setter that
# records a value and changes nothing reads exactly like one that works.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $failed = 0

    # Set both away from their defaults; each echo names the new state.
    $e = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF'
    $ok = ($e -match 'XLL DEPTH=OFF')
    Write-TestCase 'xll-set-echoes' -Pass:$ok -Fail:(-not $ok) -Detail $e
    if (-not $ok) { $failed++ }

    $e = Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'TOP'
    $ok = ($e -match 'VBA DEPTH=TOP')
    Write-TestCase 'vba-set-echoes' -Pass:$ok -Fail:(-not $ok) -Detail $e
    if (-not $ok) { $failed++ }

    # The read-back ANSWERS rather than logging: XRayStatus wrote a status
    # line to the action log, and removing it removed that line too.
    $xd = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    $vd = [string](Get-XRayTraceParam $sx 'VBA' 'DEPTH')
    $ok = ($xd -eq 'OFF' -and $vd -eq 'TOP')
    Write-TestCase 'getparam-reports-both-sources' -Pass:$ok -Fail:(-not $ok) -Detail "XLL DEPTH=$xd VBA DEPTH=$vd"
    if (-not $ok) { $failed++ }

    # TOP IS DISTINGUISHABLE FROM ALL. This is what the removed surface could
    # not do -- it reported "ON" for both -- so it is asserted explicitly here
    # rather than left implied by the case above.
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'TOP')
    $top = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'ALL')
    $all = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    $ok = ($top -eq 'TOP' -and $all -eq 'ALL')
    Write-TestCase 'top-and-all-are-distinguishable' -Pass:$ok -Fail:(-not $ok) -Detail "TOP read '$top', ALL read '$all'"
    if (-not $ok) { $failed++ }

    # Off-vocabulary: refused, and the value is UNCHANGED afterwards.
    $e = Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'BANANAS'
    $after = [string](Get-XRayTraceParam $sx 'VBA' 'DEPTH')
    $ok = ($e -match '#Err' -and $after -eq 'TOP')
    Write-TestCase 'bad-value-refused-nothing-changed' -Pass:$ok -Fail:(-not $ok) -Detail "$e / then VBA DEPTH=$after"
    if (-not $ok) { $failed++ }

    # An off-vocabulary NAME is refused the same way.
    $e = Set-XRayTraceParam $sx 'XLL' 'BANANAS' 'ALL'
    $ok = ($e -match '#Err')
    Write-TestCase 'bad-name-refused' -Pass:$ok -Fail:(-not $ok) -Detail $e
    if (-not $ok) { $failed++ }

    # BUFFERSIZE and BUFFERWHENFULL are SYMMETRIC and source-less: the getter
    # returns the value in the same grammar the setter accepts, so a cell can
    # read it and set it back. The live "did this run lose anything" counts
    # belong on the status summary, not the parameter echo.
    #
    # BUFFERSIZE takes an optional unit: bare or M/MB is megabytes, K/KB is
    # kilobytes. The getter reports the canonical form (MB when it divides
    # evenly, else KB).
    [void](Set-XRayTraceParam $sx 'BUFFERSIZE' '32')          # bare number = MB
    $b = [string](Get-XRayTraceParam $sx 'BUFFERSIZE')
    $ok = ($b -eq '32MB')
    Write-TestCase 'buffer-bare-number-is-mb' -Pass:$ok -Fail:(-not $ok) -Detail "BUFFERSIZE read '$b' (expected 32MB)"
    if (-not $ok) { $failed++ }

    [void](Set-XRayTraceParam $sx 'BUFFERSIZE' '512K')        # KB suffix, sub-MB ring
    $b = [string](Get-XRayTraceParam $sx 'BUFFERSIZE')
    $ok = ($b -eq '512KB')
    Write-TestCase 'buffer-kb-suffix-round-trips' -Pass:$ok -Fail:(-not $ok) -Detail "BUFFERSIZE read '$b' (expected 512KB)"
    if (-not $ok) { $failed++ }

    # A ring below the 16 KB floor is refused, value unchanged (still 512KB).
    $e = Set-XRayTraceParam $sx 'BUFFERSIZE' '4K'
    $after = [string](Get-XRayTraceParam $sx 'BUFFERSIZE')
    $ok = ($e -match '#Err' -and $after -eq '512KB')
    Write-TestCase 'buffer-below-floor-refused' -Pass:$ok -Fail:(-not $ok) -Detail "$e / then BUFFERSIZE=$after"
    if (-not $ok) { $failed++ }

    $e = Set-XRayTraceParam $sx 'BUFFERWHENFULL' 'PAUSE'
    $w = [string](Get-XRayTraceParam $sx 'BUFFERWHENFULL')
    $ok = ($e -match 'PAUSE' -and $w -eq 'PAUSE')
    Write-TestCase 'whenfull-get-returns-what-was-set' -Pass:$ok -Fail:(-not $ok) -Detail "echo '$e' / BUFFERWHENFULL read '$w'"
    if (-not $ok) { $failed++ }

    # Off-vocabulary BUFFERWHENFULL is refused, value unchanged (still PAUSE).
    $e = Set-XRayTraceParam $sx 'BUFFERWHENFULL' 'BANANAS'
    $after = [string](Get-XRayTraceParam $sx 'BUFFERWHENFULL')
    $ok = ($e -match '#Err' -and $after -eq 'PAUSE')
    Write-TestCase 'bad-whenfull-refused-nothing-changed' -Pass:$ok -Fail:(-not $ok) -Detail "$e / then BUFFERWHENFULL=$after"
    if (-not $ok) { $failed++ }

    if ($failed) { Complete-Test -Fail -Detail "$failed echo/query case(s) failed" }
    Complete-Test -Pass
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
