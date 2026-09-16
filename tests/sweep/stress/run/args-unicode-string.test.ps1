$case = @{ Name='args-unicode-string'
     Modules=@{
       'M'=@'
Public Sub Uni(ByVal s As String)
    Dim z As Long
    z = Len(s)
End Sub
Public Sub Go()
    Uni ChrW(937) & ChrW(8364) & ChrW(26085) & "-tail"
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Uni' } | Select-Object -First 1
        if (-not $e) { return "no row for Uni" }
        if ($e.typetext -ne 'String') { return "expected String, got [$($e.typetext)]" }
        $null }
     Why='non-ASCII argument text must not corrupt the row or its escaping' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
