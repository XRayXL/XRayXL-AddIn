$case = @{ Name='args-byref-mutated-in-body'
     Modules=@{
       'M'=@'
Public Sub Mutate(ByRef v As Long)
    v = 9999999
End Sub
Public Sub Go()
    Dim v As Long
    v = 1122867
    Mutate v
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Mutate' } | Select-Object -First 1
        if (-not $e) { return "no row for Mutate" }
        if ($e.args -notmatch '1122867') { return "entry-time value not captured: [$($e.args)]" }
        if ($e.args -match '9999999')    { return "row shows the POST-call value: [$($e.args)]" }
        $null }
     Why='arguments are captured at ENTRY, so a ByRef the body rewrites must not change the row' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
