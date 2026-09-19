$case = @{ Name='args-array-10000'
     Modules=@{
       'M'=@'
Public Sub Arr(ByRef a() As Long)
    Dim z As Long
    z = a(1)
End Sub
Public Sub Go()
    Dim a(1 To 10000) As Long
    Dim i As Long
    For i = 1 To 10000
        a(i) = i
    Next i
    Arr a
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Arr' } | Select-Object -First 1
        if (-not $e) { return "no row for Arr" }
        if ($e.args -notmatch '1\.\.10000')     { return "bounds not reported: [$($e.args)]" }
        # Arrays are not truncated: every element is in the row, the last one closing it.
        if ($e.args -notmatch '\{1,2,3,') { return "the first elements are missing: [$($e.args)]" }
        if ($e.args -notmatch ',9999,10000\}') { return "the array does not end at its last element" }
        if ($e.args -match 'shown')        { return "the array was cut: [$($e.args)]" }
        $null }
     Why='a 10,000-element array: the bounds are reported and every element is written, none cut' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
