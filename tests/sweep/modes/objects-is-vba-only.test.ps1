# OBJECTS is a VBA setting: an XLL argument has no object model to describe. So it is refused for XLL,
# and for an omitted Source, which means both. Also pins `#Err - ` refusals and case-insensitivity.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx

    # ---- the refusals -----------------------------------------------------
    $e = Set-XRayTraceParam $sx 'XLL' 'OBJECTS' 'FALSE'
    Check 'xll-objects-refused' ($e -match '^#Err - ') $e
    Check 'xll-objects-says-vba-only' ($e -match 'only available for VBA') $e

    $e = Set-XRayTraceParam $sx $null 'OBJECTS' 'FALSE'
    Check 'omitted-source-objects-refused' ($e -match '^#Err - ') $e
    Check 'omitted-source-says-vba-only' ($e -match 'only available for VBA') $e

    # ...and the refusal CHANGED NOTHING: VBA still holds its default.
    $v = [string](Get-XRayTraceParam $sx 'VBA' 'OBJECTS')
    Check 'refusal-left-vba-alone' ($v -eq 'TRUE') "VBA OBJECTS='$v'"

    # ---- the one source that does honour it -------------------------------
    $e = Set-XRayTraceParam $sx 'VBA' 'OBJECTS' 'FALSE'
    Check 'vba-objects-accepted' ($e -notmatch '#Err') $e
    $v = [string](Get-XRayTraceParam $sx 'VBA' 'OBJECTS')
    Check 'vba-objects-took-effect' ($v -eq 'FALSE') "VBA OBJECTS='$v'"

    # ---- a read answers rather than errors, but not TRUE/FALSE for a setting nothing reads ----
    $v = [string](Get-XRayTraceParam $sx 'XLL' 'OBJECTS')
    Check 'get-xll-objects-is-not-a-value' ($v -notmatch '^(TRUE|FALSE)$') "XLL OBJECTS='$v'"
    Check 'get-xll-objects-says-vba-only' ($v -match 'VBA only') "XLL OBJECTS='$v'"

    # ---- case-insensitivity, asserted on a setting that really changes ----
    $e = Set-XRayTraceParam $sx 'vba' 'objects' 'true'
    Check 'lower-case-accepted' ($e -notmatch '#Err') $e
    $v = [string](Get-XRayTraceParam $sx 'VBA' 'OBJECTS')
    Check 'lower-case-took-effect' ($v -eq 'TRUE') "VBA OBJECTS='$v'"

    $e = Set-XRayTraceParam $sx 'xll' 'depth' 'top'
    Check 'lower-case-depth-accepted' ($e -notmatch '#Err') $e
    $v = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    Check 'lower-case-depth-took-effect' ($v -eq 'TOP') "XLL DEPTH='$v'"

    # ---- and a mixed-case refusal still refuses, by the same route --------
    $e = Set-XRayTraceParam $sx 'Xll' 'Objects' 'False'
    Check 'mixed-case-objects-still-refused' ($e -match '^#Err - ') $e

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'OBJECTS is VBA only; refusals say #Err - and words are case-insensitive'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
