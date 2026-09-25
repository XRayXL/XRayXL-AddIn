# The VBA editor runs a macro through a hidden procedure of its own, `_ImmedProc`. The trace leaves
# it out, so the macro reads as the top of its chain with `editor` as its caller, as its callees do.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$code = @'
Sub Outer()
    Inner
End Sub

Sub Inner()
    ThisWorkbook.Worksheets(1).Range("E1").Value = "ran"
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')
    New-XRayMacroBook $sx 'EditorRun' -Components @(@{ Kind = 1; Name = 'M'; Code = $code })
    $book = Get-XRayMacroBook 'EditorRun'

    # The cursor inside Outer, then the editor's own Run command, as F5 would press it.
    $module = $book.Book.VBProject.VBComponents.Item('M').CodeModule
    $at = $module.ProcStartLine('Outer', 0) + 1
    $module.CodePane.Show()
    $module.CodePane.SetSelection($at, 1, $at, 1)
    $run = $app.VBE.CommandBars.FindControl([Type]::Missing, 186)    # Run Sub/UserForm
    if (-not $run) { Complete-Test -Skip -Detail "the editor has no Run Sub/UserForm command to press" }

    $s = Invoke-XRayArmedSession $sx -Body { $run.Execute() }
    if ([string]$book.Sheet.Range('E1').Value2 -ne 'ran') {
        Complete-Test -Skip -Detail "the editor's Run command did not run the macro when pressed from automation"
    }
    $rows = @(Read-TraceRows $sx.ProcId | Where-Object { $_.source -eq 'VBA' -and $_.kind -eq 'entry' })

    Check 'the-wrapper-is-left-out' (-not ($rows | Where-Object { $_.function -eq '_ImmedProc' })) `
          ("entries: " + (($rows | ForEach-Object { "$($_.function)@$($_.depth)" }) -join ' '))
    $outer = $rows | Where-Object { $_.function -eq 'Outer' } | Select-Object -First 1
    $inner = $rows | Where-Object { $_.function -eq 'Inner' } | Select-Object -First 1
    Check 'the-macro-is-the-top' ($outer -and $outer.depth -eq '1' -and $outer.parent -eq '0') `
          "Outer depth '$($outer.depth)' parent '$($outer.parent)'"
    Check 'its-caller-is-the-editor' ($outer.caller -eq 'editor' -and -not $outer.callerref) `
          "caller '$($outer.caller)' callerref '$($outer.callerref)'"
    Check 'its-callee-hangs-from-it-and-inherits-the-caller' `
          ($inner -and $inner.depth -eq '2' -and $inner.parent -eq $outer.span -and $inner.caller -eq 'editor') `
          "Inner depth '$($inner.depth)' parent '$($inner.parent)' caller '$($inner.caller)'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("entries: " + (($rows | ForEach-Object { "$($_.function)@$($_.depth):$($_.caller)" }) -join ' '))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
