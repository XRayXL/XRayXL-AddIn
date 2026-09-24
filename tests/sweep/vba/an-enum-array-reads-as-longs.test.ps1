# An array of an Enum reads as the Longs it holds. Its descriptor says VT_USERDEFINED, which
# names no element type, and carries no Enum name, so none is claimed; TypeName says Long().
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
    $args1 = Remove-ArgAddress (ArgsOf $s.Rows 'TakeColours')

    Check 'an-enum-array-reads-as-longs' ($args1 -eq 'a1:Ref&=Long[1..3]{1,2,7}') "args: $args1"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "args: $args1"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
