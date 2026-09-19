# A WORKBOOK THAT WAS NEVER SAVED.
#
# VBE7 keeps the project's file name where the module column reads the workbook from, and a book
# that was never saved has none: the field holds an internal id (ten hex digits). The id is the
# only handle there is, so it is kept -- but marked, because "006cd206e1" otherwise reads as a
# workbook a user could go and look for.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub UnsavedProbe()
    UnsavedHelper 1
End Sub

Private Sub UnsavedHelper(ByVal n As Long)
    Dim x As Long
    x = n + 1
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $booksRef = $app.Workbooks
    $wb = $booksRef.Add()
    try { $wb.EnableAutoRecover = $false } catch {}
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.Name = 'MUnsaved'
    $comp.CodeModule.AddFromString($moduleCode)

    $s = Invoke-XRayArmedSession $sx -Body { $app.Run("'" + $wb.Name + "'!UnsavedProbe") | Out-Null }
    $rows = @(Read-TraceRows $sx.ProcId | Where-Object { $_.function -eq 'UnsavedProbe' -or $_.function -eq 'UnsavedHelper' })
    $wb.Close($false)

    $modules = @($rows | ForEach-Object { $_.module } | Select-Object -Unique)
    $one = [string](@($modules) | Select-Object -First 1)

    Check 'it-was-traced' ([bool]($rows.Count -ge 2)) "rows: $($rows.Count)"
    Check 'the-book-is-marked-unsaved' ([bool]($one -match '^\[unsaved:[0-9a-f]{6,}\]MUnsaved$')) "module: '$one'"
    Check 'every-row-says-the-same' ([bool]($modules.Count -eq 1)) ($modules -join ' | ')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "module: $one"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
