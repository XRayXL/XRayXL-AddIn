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
        if ($e.args -notmatch 'of 10000 shown') { return "truncation not declared: [$($e.args)]" }
        $null }
     Why='a 10,000-element array: the true count must be reported and the render bounded' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
