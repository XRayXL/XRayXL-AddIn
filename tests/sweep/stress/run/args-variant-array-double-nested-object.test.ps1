$case = @{ Name='args-variant-array-double-nested-object'
     Modules=@{
       'M'=@'
' One ByVal Variant argument holding an array of a Double, another array, an object and a Boolean.
Public Function TakeV(ByVal v As Variant) As Long
    TakeV = UBound(v)
End Function

Public Sub Go()
    Dim c As New Collection
    c.Add 1
    Dim n As Long
    n = TakeV(Array(1.5, Array(2, 3), c, True))
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq 'TakeV' }) | Select-Object -First 1
        if (-not $e) { return 'TakeV was not traced' }
        # Each element by its own rule: a bare Double, a nested array with its bounds,
        # the object by class and address, and a Boolean spelt as Excel spells it.
        $want = '^a1:Variant=Variant\[0\.\.3\]\{1\.5,Variant\[0\.\.1\]\{Integer\(2\),Integer\(3\)\},Collection@0x[0-9A-Fa-f]+,TRUE\}$'
        if ([string]$e.args -cnotmatch $want) { return "args [$($e.args)] do not match $want" }
        $null }
     Why='a Variant array of a Double, a nested array, an object and a Boolean in one argument: every element renders by its own rule inside the one args grammar' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
