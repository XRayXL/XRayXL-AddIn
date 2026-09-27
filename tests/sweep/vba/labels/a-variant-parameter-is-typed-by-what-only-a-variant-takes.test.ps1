# A ByRef Variant parameter the body only indexes, treats as an array, or hands to an object's
# method has no typed load of its own. What it is pushed straight into says what it is: an
# operation only a Variant takes (CRefVarAry, the VarIndex family), or 752, which pushes a ByRef
# Variant's value whole for a method's argument. An array parameter indexed the same way stays an
# array. Diagnostics are on (suite.psd1), so each signature names the opcode that typed it.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

$moduleCode = @'
Public gN As Double, gV As Variant

Public Sub UseLBound(v As Variant)
    gN = LBound(v)
End Sub
Public Sub UseUBound2(v As Variant)
    gN = UBound(v, 1)
End Sub
Public Sub UseIndexLoad(v As Variant)
    gV = v(1)
End Sub
Public Sub UseIndexStore(v As Variant)
    v(1) = 42
End Sub
Public Sub UseIndexCom(v As Variant)
    gN = ThisWorkbook.Worksheets(1).Cells(v(0), 1).Row
End Sub
Public Sub UseCom(r As Variant)
    gN = ThisWorkbook.Worksheets(1).Cells(r, 1).Row
End Sub
Public Sub UseIndexPass(v As Variant)
    TakeRef v(0)
End Sub
Public Sub UseWhole(v As Variant)
    TakeVal v
End Sub
Public Sub TakeRef(x As Variant)
    gV = x
End Sub
Public Sub TakeVal(ByVal x As Variant)
    gV = x
End Sub
Public Sub UseArr(a() As Long)
    gN = a(1)
End Sub

Public Sub Drive()
    Dim v As Variant, w As Variant, r As Variant, a(0 To 2) As Long
    v = Array(10, 20, 30)
    w = Array(2, 3)
    r = 4
    a(1) = 7
    UseLBound v
    UseUBound2 v
    UseIndexLoad v
    UseIndexPass v
    UseWhole v
    UseArr a
    UseIndexCom w
    UseCom r
    UseIndexStore v
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'VariantUse' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    function Param1([string]$fn) {
        $sig = ((SigOf $rows $fn) -split ',')[0] -replace '\|.*$', ''
        $val = [regex]::Match([string](Remove-ArgAddress (ArgsOf $rows $fn)), '^a1:[^=]*=(.*)$').Groups[1].Value
        [pscustomobject]@{ Type = ($sig -replace '#.*$', ''); Op = $(if ($sig -match '#(\d+)') { [int]$Matches[1] } else { 0 }); Value = $val }
    }
    function Expect([string]$fn, [string]$type, [int[]]$ops, [string]$value) {
        $p = Param1 $fn
        Write-Output ('  {0,-14} {1,-24} {2}' -f $fn, (SigOf $rows $fn), (Remove-ArgAddress (ArgsOf $rows $fn)))
        $ok = $p.Type -ceq $type -and ($ops.Count -eq 0 -or $ops -contains $p.Op) -and $p.Value -ceq $value
        Check $fn $ok ("[$(SigOf $rows $fn)] a1=$($p.Value); want $type#$($ops -join '/') = $value")
    }

    $three = 'Variant[0..2]{Integer(10),Integer(20),Integer(30)}'
    Write-Output ''
    Expect 'UseLBound'     'Variant&' @(437)  $three
    Expect 'UseUBound2'    'Variant&' @(437)  $three
    Expect 'UseIndexLoad'  'Variant&' @(1510) $three
    Expect 'UseIndexPass'  'Variant&' @(1605) $three
    Expect 'UseIndexCom'   'Variant&' @(1511) 'Variant[0..1]{Integer(2),Integer(3)}'
    Expect 'UseCom'        'Variant&' @(752)  'Integer(4)'
    Expect 'UseIndexStore' 'Variant&' @(1512) $three
    # passed whole to a ByVal Variant, it is loaded as one (742) before any of these
    Expect 'UseWhole'      'Variant&' @(742)  $three
    Expect 'UseArr'        'Ref&'     @()     'Long[0..2]{0,7,0}'

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("every Variant use typed; UseLBound [{0}]" -f (SigOf $rows 'UseLBound'))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
