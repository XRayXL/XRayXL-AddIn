# An empty or never-allocated array is a value, and reads as an array, not as a number.
#
#    =PA()                 an empty ParamArray: VBA passes a descriptor with no element type,
#                          no width and no data. It printed as the pointer in decimal
#    EmptyVar()            a Variant holding a Long() never allocated: the read was declined
#                          and both ret and rettype were blank
#    EmptyTyped()          a Function As Long() that never allocated its result
#    FillIt a  (a unallocated, ByRef)   a zero in a Ref& slot: a LongLong 0 or an unallocated
#                          array, which nothing tells apart, so the raw qword
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Function PA(ParamArray p() As Variant) As Long
    PA = UBound(p) - LBound(p) + 1
End Function

Public Function EmptyVar() As Variant
    Dim e() As Long
    EmptyVar = e
End Function

Public Function EmptyTyped() As Long()
    Dim e() As Long
    EmptyTyped = e
End Function

Public Sub Drive()
    Dim v As Variant
    v = EmptyVar()
    Dim t() As Long
    t = EmptyTyped()
    Dim a() As Long
    FillIt a
End Sub

Private Sub FillIt(ByRef a() As Long)
    ReDim a(1 To 2)
    a(1) = 7
End Sub
'@

function EntryOf($Rows, [string]$Fn) { @($Rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq $Fn }) | Select-Object -First 1 }
function ExitOf($Rows, [string]$Fn)  { @($Rows | Where-Object { $_.kind -eq 'exit'  -and $_.source -eq 'VBA' -and $_.function -eq $Fn }) | Select-Object -First 1 }

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    New-XRayMacroBook $sx 'EmptyArrays' @(
        @{ Kind=1; Name='M'; Code=$moduleCode }
    ) -Cells @{ 'A1' = '=PA()'; 'A2' = '=PA(1,2)' }
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    Invoke-XRayRecalc $app
    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $pa = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'PA' })
    $empty = $pa | Where-Object { $_.callerref -like '*!A1' } | Select-Object -First 1
    $two   = $pa | Where-Object { $_.callerref -like '*!A2' } | Select-Object -First 1
    Check 'an-empty-paramarray-reads-as-an-empty-array' ($empty.args -eq 'a1:Ref&=?[0..-1]{}') "args: $($empty.args)"
    Check 'a-filled-paramarray-is-unchanged' ($two.args -eq 'a1:Ref&=Variant[0..1]{1,2}') "args: $($two.args)"

    $ev = ExitOf $rows 'EmptyVar'
    Check 'an-unallocated-array-in-a-variant-reads-Long()' (($ev.ret -eq 'Long()') -and ($ev.rettype -eq 'Variant')) "ret '$($ev.ret)' rettype '$($ev.rettype)'"
    $et = ExitOf $rows 'EmptyTyped'
    Write-XRayObservation 'typed-unallocated-return' "ret '$($et.ret)' rettype '$($et.rettype)'"
    Check 'a-typed-unallocated-return-is-not-a-number' ($et.ret -notmatch '^-?\d+$') "ret '$($et.ret)' rettype '$($et.rettype)'"

    $fe = EntryOf $rows 'FillIt'; $fx = ExitOf $rows 'FillIt'
    Check 'an-unallocated-byref-array-reads-as-the-raw-qword' ($fe.args -eq 'a1:Ref&=0x0') "entry args: $($fe.args)"
    Check 'its-allocation-shows-at-the-exit' ($fx.args -eq 'a1:Ref&=Long[1..2]{7,0}') "exit args: $($fx.args)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "empty ParamArray $($empty.args); unallocated Variant ret $($ev.ret)"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
