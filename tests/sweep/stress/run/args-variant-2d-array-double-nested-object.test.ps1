$case = @{ Name='args-variant-2d-array-double-nested-object'
     Modules=@{
       'M'=@'
' A 2-D Variant array of a Double, another array, an object and a Boolean, in one ByVal Variant argument.
Public Function Take2(ByVal v As Variant) As Long
    Take2 = UBound(v, 2)
End Function

Public Sub Go()
    Dim c As New Collection
    c.Add 1
    Dim a(1 To 2, 1 To 2) As Variant
    a(1, 1) = 1.5
    a(1, 2) = Array(2, 3)
    Set a(2, 1) = c
    a(2, 2) = True
    Dim n As Long
    n = Take2(a)
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq 'Take2' }) | Select-Object -First 1
        if (-not $e) { return 'Take2 was not traced' }
        # A level per row: {a(1,1),a(1,2)},{a(2,1),a(2,2)}. The bytes lie column-major, which
        # would put the object in the first row. Array(2, 3) holds Integers, so they are named.
        $want = '^a1:Variant=Variant\[1\.\.2,1\.\.2\]\{\{1\.5,Variant\[0\.\.1\]\{Integer\(2\),Integer\(3\)\}\},\{Collection@0x[0-9A-Fa-f]+,TRUE\}\}$'
        if ([string]$e.args -cnotmatch $want) { return "args [$($e.args)] do not match $want" }
        $null }
     Why='a 2-D Variant array of a Double, a nested array, an object and a Boolean: elements read row by row, each by its own rule' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
