# A class module in a workbook opened from disk is named, like one built in the session.
#
# Every other case builds its workbook in the session. A class module loaded from a saved
# workbook carries a different marker in its parent and module entry, 0x8000000000000008 in
# place of -1, and its module entry has no back-pointer; the old check refused that chain and
# the procedures read as addresses. The committed demo 06_Objects.xlsm is opened here because
# no session has loaded it before.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$book = Join-Path $PSScriptRoot '..\..\..\dist\demo\06_Objects.xlsm'

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    if (-not (Test-Path $book)) { Complete-Test -Fail -Detail "missing: $book" }

    $wb = $app.Workbooks.Open((Resolve-Path $book).Path)
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $app.CalculateFull()
    $lossy = Stop-XRayTrace $sx
    $wb.Close($false)
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $entries = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' })
    $class = @($entries | Where-Object { $_.module -eq '[06_Objects.xlsm]Position' })
    $names = @($class | ForEach-Object { $_.function })

    Check 'the-class-members-are-named' (($names -join ',') -eq 'Class_Initialize,Qty,Value,Class_Terminate') `
          "Position members: $($names -join ',')"
    $unnamed = @($entries | Where-Object { $_.function -match '^0x[0-9A-F]+$' })
    Check 'no-procedure-reported-as-an-address' ($unnamed.Count -eq 0) `
          ("unnamed: " + (@($unnamed | ForEach-Object { $_.proc }) -join ','))
    # the negative control: the standard module in the same workbook was always named
    Check 'the-standard-module-is-named' (@($entries | Where-Object { $_.function -eq 'PositionValue' -and $_.module -eq '[06_Objects.xlsm]Objects' }).Count -eq 1) `
          "PositionValue rows: $(@($entries | Where-Object { $_.function -eq 'PositionValue' }).Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "class members from disk named: $($names -join ',')"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
