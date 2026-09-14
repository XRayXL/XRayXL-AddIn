$case = @{ Name='args-huge-string'
     Modules=@{
       'M'=@'
Public Sub Big(ByVal s As String)
    Dim z As Long
    z = Len(s)
End Sub
Public Sub Go()
    Big String(4000, "X")
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Big' } | Select-Object -First 1
        if (-not $e) { return "no row for Big" }
        if ($e.typetext -ne 'String') { return "expected String, got [$($e.typetext)]" }
        if ($e.args.Length -gt 600) { return "args field is $($e.args.Length) chars -- unbounded" }
        $null }
     Why='a 4000-character argument: the decoder must bound the row but still name the type' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
