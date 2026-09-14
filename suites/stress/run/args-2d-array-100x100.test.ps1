$case = @{ Name='args-2d-array-100x100'
     Modules=@{
       'M'=@'
Public Sub Arr2(ByRef m() As Double)
    Dim z As Double
    z = m(1, 1)
End Sub
Public Sub Go()
    Dim m(1 To 100, 1 To 100) As Double
    m(1, 1) = 1.5
    Arr2 m
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Arr2' } | Select-Object -First 1
        if (-not $e) { return "no row for Arr2" }
        if ($e.args -notmatch '1\.\.100,1\.\.100') { return "2-D bounds wrong: [$($e.args)]" }
        $null }
     Why='a 10,000-cell two-dimensional array: both bounds must survive, in declaration order' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
