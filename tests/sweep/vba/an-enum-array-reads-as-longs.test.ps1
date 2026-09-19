# An array of an Enum reads as the Longs it holds.
#
# VBA writes VT_USERDEFINED (29) in the descriptor's vartype slot, which names no readable
# element type; the elements are four-byte Longs, and VBA's own TypeName says Long(). The
# declared Enum name is not in the descriptor, so it is not claimed.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Enum Colour
    Red = 1
    Green = 2
End Enum

Public Sub Drive()
    Dim c(1 To 3) As Colour
    c(1) = Red
    c(2) = Green
    c(3) = 7
    TakeColours c
End Sub

Private Sub TakeColours(ByRef a() As Colour)
    Dim n As Long
    n = UBound(a)
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    New-XRayMacroBook $sx 'EnumArray' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { $app.Run($leaf + '!Drive') | Out-Null }
    $args1 = ArgsOf $s.Rows 'TakeColours'

    Check 'an-enum-array-reads-as-longs' ($args1 -eq 'a1:Ref&=Long[1..3]{1,2,7}') "args: $args1"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "args: $args1"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
