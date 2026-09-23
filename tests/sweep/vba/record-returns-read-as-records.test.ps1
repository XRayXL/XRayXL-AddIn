# A Function returning a user-defined Type leaves through exit 1494, and the caller's result
# buffer arrives as a hidden first argument. Two things must hold:
#
#   1. The hidden argument is not a parameter: args, typetext and argcount name only what the
#      source declares. It once read `a1:?unseen=0x... a2:Long=4`, every parameter one slot late.
#   2. The return is the record, `udt@0x...` with rettype Udt, as a record argument reads. It once
#      read a Double, 4.5: with the exit unmapped, the store scan took a field for the result.
#
# A record of 1, 2, 4 or 8 bytes comes back in a register instead, through exit 1495, with no
# hidden argument. Unmapped, it read as its one field: `Long 4`, `String "seven"`.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$src = @'
Option Explicit

Public Type RRet
    s As String
    n As Long
    d As Double
End Type

Private Type RByte: b As Byte: End Type
Private Type RLong: n As Long: End Type
Private Type RTwo: a As Long: b As Long: End Type
Private Type RStr: s As String: End Type

Private Function MakeRet(ByVal k As Long) As RRet
    MakeRet.s = "r" & k
    MakeRet.n = k * 10
    MakeRet.d = k + 0.5
End Function

Private Function MakeRet2(ByVal a As Long, ByVal b As String) As RRet
    MakeRet2.n = a
    MakeRet2.s = b
End Function

Private Function SmallByte(ByVal k As Long, ByVal tag As String) As RByte
    SmallByte.b = k
End Function
Private Function SmallLong(ByVal k As Long, ByVal tag As String) As RLong
    SmallLong.n = k
End Function
Private Function SmallTwo(ByVal k As Long, ByVal tag As String) As RTwo
    SmallTwo.b = k
End Function
Private Function SmallStr(ByVal k As Long, ByVal tag As String) As RStr
    SmallStr.s = tag & k
End Function

Public Sub Drive()
    Dim r As RRet, q As RRet, sb As RByte, sl As RLong, st As RTwo, ss As RStr
    r = MakeRet(4)
    q = MakeRet2(7, "seven")
    sb = SmallByte(1, "one")
    sl = SmallLong(2, "two")
    st = SmallTwo(3, "three")
    ss = SmallStr(4, "four")
    If r.n <> 40 Or q.s <> "seven" Or sb.b <> 1 Or sl.n <> 2 Or st.b <> 3 Or ss.s <> "four4" Then Err.Raise 5
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'udtret' @( @{ Kind = 1; Name = 'MUdt'; Code = $src } ) -Leaf 'UdtRet.xlsm'
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    [void]$app.Run('MUdt.Drive')
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))

    function Row($fn, $kind) { @($rows | Where-Object { $_.function -eq $fn -and $_.kind -eq $kind }) | Select-Object -First 1 }

    $e1 = Row 'MakeRet' 'entry'; $e2 = Row 'MakeRet2' 'entry'
    Check 'one-parameter-is-one' ($e1 -and $e1.argcount -eq '1' -and $e1.typetext -eq 'Long' -and $e1.args -eq 'a1:Long=4') `
          "MakeRet argcount='$($e1.argcount)' typetext='$($e1.typetext)' args='$($e1.args)'"
    Check 'two-parameters-are-two' ($e2 -and $e2.argcount -eq '2' -and $e2.typetext -eq 'Long,String' -and
                                    $e2.args -match '^a1:Long=7 ' -and $e2.args -match '"seven"') `
          "MakeRet2 argcount='$($e2.argcount)' typetext='$($e2.typetext)' args='$($e2.args)'"

    foreach ($fn in 'SmallByte', 'SmallLong', 'SmallTwo', 'SmallStr') {
        $e = Row $fn 'entry'
        Check "$fn-has-no-hidden-argument" ($e -and $e.argcount -eq '2' -and $e.typetext -eq 'Long,String' -and $e.args -match '^a1:Long=\d ') `
              "$fn argcount='$($e.argcount)' typetext='$($e.typetext)' args='$($e.args)'"
    }

    foreach ($fn in 'MakeRet', 'MakeRet2', 'SmallByte', 'SmallLong', 'SmallTwo', 'SmallStr') {
        $x = Row $fn 'exit'
        Check "$fn-returns-a-record" ($x -and $x.rettype -eq 'Udt' -and $x.ret -match '^udt@0x[0-9A-F]+$') `
              "$fn ret='$($x.ret)' rettype='$($x.rettype)'"
    }

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'a record-returning Function, large or small, has only its declared parameters and returns udt@'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
