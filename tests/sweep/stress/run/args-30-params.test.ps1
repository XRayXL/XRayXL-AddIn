$case = @{ Name='args-30-params'
     Modules=@{
       'M'=@'
Public Sub Wide( _
    ByVal a00 As Long, ByVal a01 As Long, ByVal a02 As Long, ByVal a03 As Long, ByVal a04 As Long, _
    ByVal a05 As Long, ByVal a06 As Long, ByVal a07 As Long, ByVal a08 As Long, ByVal a09 As Long, _
    ByVal a10 As Long, ByVal a11 As Long, ByVal a12 As Long, ByVal a13 As Long, ByVal a14 As Long, _
    ByVal a15 As Long, ByVal a16 As Long, ByVal a17 As Long, ByVal a18 As Long, ByVal a19 As Long, _
    ByVal a20 As Long, ByVal a21 As Long, ByVal a22 As Long, ByVal a23 As Long, ByVal a24 As Long, _
    ByVal a25 As Long, ByVal a26 As Long, ByVal a27 As Long, ByVal a28 As Long, ByVal a29 As Long)
    Dim z As Long
    z = a00
    z = a59
End Sub
Public Sub Go()
    Wide _
    1, 2, 3, 4, 5, _
    6, 7, 8, 9, 10, _
    11, 12, 13, 14, 15, _
    16, 17, 18, 19, 20, _
    21, 22, 23, 24, 25, _
    26, 27, 28, 29, 30
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Wide' } | Select-Object -First 1
        if (-not $e) { return "no row for Wide" }
        if ([int]$e.argcount -ne 30) { return "expected 30 arg slots, got $($e.argcount)" }
        $null }
     Why='a 30-parameter signature: argSz and the frame region stretched. VBA caps a logical line near 1023 chars AND line-continuations near 25, so 60 params will not COMPILE however they are wrapped' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
