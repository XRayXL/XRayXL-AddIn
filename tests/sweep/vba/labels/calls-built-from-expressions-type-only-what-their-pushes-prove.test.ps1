# Calls whose arguments are expressions, nested calls, named and omitted arguments, record members,
# array elements and temporaries, into procedures whose own code never touches their parameters.
# Read back from the call, each instruction that is one push is the next slot, from slot 1; the
# first that is not ends what the call can say. A local passed by address says nothing: a record's
# first member is stored at the record's own offset. Every case states, for each parameter, the
# declared type and the value passed, or that it stays untyped.
#
#   T=V   typed as T, with value V         ?   must stay untyped
# A ByVal String is typed by its own code even unused (663), so its caller is never asked. An
# Integer literal names nothing: it also fills a Byte parameter unconverted.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

# Parameters their own code never touches: only a caller can type them.
$sinks = @(
    @{ N='K01'; D='Sub K01(ByVal a As Long, ByVal b As Double, ByVal c As String)'; C='K01 p, FD(1) * 2, s'; W=@('Long=3','?','String="sv"') }
    @{ N='K02'; D='Sub K02(ByVal a As Long, ByVal b As Double, ByVal c As String)'; C='K02 7, 2.5, "lit"'; W=@('Long=7','Double=2.5','String="lit"') }
    # a call evaluated after other arguments goes through a temporary and a typed load
    @{ N='K03'; D='Sub K03(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal e As Long)'; C='K03 FL(1), 2, 3, 4, 5'; W=@('Long=2','Long=2','Long=3','Long=4','Long=5') }
    @{ N='K04'; D='Sub K04(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal e As Long)'; C='K04 1, 2, 3, 4, FL(5)'; W=@('Long=1','Long=2','Long=3','Long=4','?') }
    @{ N='K05'; D='Function K05(ByVal a As Long, ByVal b As String) As Variant'; C='v = K05(p, s)'; W=@('Long=3','String="sv"') }
    @{ N='K06'; D='Function K06(ByVal a As Long, ByVal b As Long) As Rec'; C='r2 = K06(p + 1, p)'; W=@('?','?') }
    # a ByVal Variant's own exit clears it, so its own code types it, three slots wide
    @{ N='K07'; D='Sub K07(ByVal v As Variant, ByVal a As Long, ByVal b As String)'; C='K07 p, p, s'; W=@('Variant=Long(3)','?','String="sv"') }
    @{ N='K08'; D='Sub K08(ByVal a As Long, ByVal v As Variant, ByVal b As String)'; C='K08 p, p, s'; W=@('Long=3','Variant=Long(3)','String="sv"') }
    @{ N='K09'; D='Sub K09(Optional ByVal a As Long = 9, Optional ByVal b As String = "d", Optional ByVal c As Long = 4)'; C='K09 , , 3'; W=@('Long=9','String="d"','Long=3') }
    @{ N='K10'; D='Sub K10(Optional ByVal a As Long = 9, Optional ByVal b As String = "d", Optional ByVal c As Long = 4)'; C='K10 c:=FL(2), a:=5'; W=@('Long=5','String="d"','?') }
    @{ N='K11'; D='Sub K11(a As Long, b As String, rr As Rec, ar() As Long)'; C='K11 n, t, r, arr'; W=@('?','?','?','?') }
    @{ N='K12'; D='Sub K12(a As Long, b As String, rr As Rec, ar() As Long)'; C='K12 rp, rs, r, arr'; W=@('Long&=8','String&="rs"','?','?') }
    @{ N='K13'; D='Sub K13(ByVal b As Boolean, ByVal i As Integer, ByVal c As Currency, ByVal g As Single, ByVal q As LongLong, ByVal d As Date)'; C='K13 True, 3, 2.5@, 1.5!, 5^, #1/2/2020#'; W=@('?','?','Currency=2.5000','Single=1.5','LongLong=5','Double=43832') }
    @{ N='K14'; D='Property Let K14(ByVal k As Long, ByVal v As String)'; C='K14(p) = s & "!"'; W=@('Long=3','String="sv!"') }
    @{ N='K15'; D='Property Let K15(ByVal k As Long, ByVal v As String)'; C='K15(FL(p)) = s'; W=@('?','String="sv"') }
    @{ N='K16'; D='Sub K16(ByVal a As Long, ByVal b As String, ByVal c As Double)'; C='K16 FL(FL(p)), FS(FS(s)), FD(FD(q))'; W=@('Long=5','String="sv!!"','?') }
    @{ N='K17'; D='Sub K17(ByVal a As Long, ByVal b As String, ByVal c As Long)'; C='K17 p, s, FL(p) + FL(p) * 2'; W=@('Long=3','String="sv"','?') }
    @{ N='K18'; D='Sub K18(ByVal a As Long, ByVal b As Long)'; C='K18 IIf(p > 0, 1, 2), p'; W=@('?','?') }
    @{ N='K19'; D='Sub K19(ByVal a As Long, ByVal b As Long)'; C='If p > 0 Then K19 p, 1 Else K19 2, p'; W=@('Long=3','Long=1') }
    @{ N='K20'; D='Function K20(ByVal a As Long, ByVal b As String) As Boolean'; C='If FL(p) > 0 And K20(p, s) Then gSink = 1'; W=@('Long=3','String="sv"') }
    @{ N='K21'; D='Function K21(ByVal a As Long, ByVal b As Long) As Long'; C='For i = FL(0) To K21(p, 2): gSink = gSink + i: Next i'; W=@('Long=3','Long=2') }
    @{ N='K22'; D='Function K22(ByVal a As Long, ByVal b As String) As Long'; C="Select Case K22(p, s)`r`n    Case Is > 0: gSink = 2`r`n    End Select"; W=@('Long=3','String="sv"') }
    @{ N='K23'; D='Sub K23(ByVal a As Long, ByVal b As Double)'; C="With r`r`n    K23 .N, .D`r`n    End With"; W=@('?','?') }
    @{ N='K24'; D='Sub K24(ByVal a As Long, ByVal b As Double)'; C='K24 r.N, r.D'; W=@('Long=6','Double=4.5') }
    @{ N='K25'; D='Sub K25(ByVal a As Long, ByVal b As Long)'; C='K25 arr(1), p'; W=@('?','?') }
    @{ N='K26'; D='Sub K26(a As Long, ByVal b As String)'; C='K26 p + 1, s'; W=@('?','String="sv"') }
    @{ N='K27'; D='Sub K27(ByVal a As Long, Optional ByVal b As String = "", Optional ByVal c As Long = 1)'; C='K27 b:=FS(s), a:=p'; W=@('Long=3','String="sv!"','?') }
    @{ N='K28'; D='Sub K28(ByVal a As Long, ByVal b As Long, ByVal c As String)'; C='K28 p, rp, rs'; W=@('Long=3','Long=8','String="rs"') }
    @{ N='K29'; D='Sub K29(ByVal a As Long, ByVal b As Double, ByVal c As String)'; C='M2.X29 p, q, s'; W=@() }
    @{ N='K30'; D='Sub K30(ByVal a As Long, ByVal b As Long)'; C='K30 p, K30f(K30f(p, 1), FL(p))'; W=@('Long=3','?') }
    # an omitted Byte's default is an Integer literal, unconverted, so that names neither; a Byte
    # literal written out is converted, which ends the read
    @{ N='K32'; D='Sub K32(ByVal n As Long, Optional ByVal b As Byte = 7)'; C='K32 p'; W=@('Long=3','?') }
    @{ N='K33'; D='Sub K33(ByVal b As Byte, ByVal n As Long)'; C='K33 5, p'; W=@('?','?') }
    @{ N='K31'; D='Function K31(ByVal a As Long, ByVal b As Long) As Long'; C='gSink = K31(p, 1) + K31(2, p) * K31(FL(p), p)'; W=@('Long=3','Long=1') }
)
# Parameters their own code only passes on: only the procedure they go to can type them.
$passers = @(
    @{ N='P01'; D='Sub P01(x As Long, y As Double, z As String)'; B='U3 x, y, z'; A='21, 0.5, "pz"'; W=@('Long&=21','Double&=0.5','String&="pz"') }
    @{ N='P02'; D='Sub P02(x As Long, z As String)'; B='U2 x, FD(2) * 3, z'; A='22, "pq"'; W=@('Long&=22','?') }
    @{ N='P03'; D='Sub P03(x As Long, z As String)'; B="U2 x, FD(2) * 3, z`r`n    U2 x, 2.5, z"; A='23, "pr"'; W=@('Long&=23','String&="pr"') }
    @{ N='P04'; D='Sub P04(y As Double, z As String)'; B='U3 FL(1), y, z'; A='0.25, "ps"'; W=@('?','?') }
    @{ N='P05'; D='Sub P05(x As Long, z As String)'; B='U3 c:=z, a:=x, b:=FD(1)'; A='25, "pt"'; W=@('Long&=25','?') }
    @{ N='P06'; D='Sub P06(ByVal k As Long)'; B='U3 k, 1.5, "a"'; A='26'; W=@('Long=26') }
    @{ N='P07'; D='Sub P07(x As Long, y As Double, z As String)'; B='M2.U3M x, y, z'; A='27, 0.75, "pu"'; W=@('Long&=27','Double&=0.75','String&="pu"') }
    @{ N='P08'; D='Sub P08(x As Long, z As String)'; B='U2 x, UF(x, 1.5) * FD(UF(x, 2)), z'; A='28, "pv"'; W=@('Long&=28','?') }
    @{ N='P09'; D='Sub P09(x As Long, z As String)'; B='If FL(1) > 0 Then U2 x, 1, z Else U2 x, FD(1) * 2, z'; A='29, "pw"'; W=@('Long&=29','?') }
)

$m = New-Object System.Text.StringBuilder
[void]$m.AppendLine('Option Explicit')
[void]$m.AppendLine('Public gSink As Double')
[void]$m.AppendLine('Public Type Rec')
[void]$m.AppendLine('    D As Double')
[void]$m.AppendLine('    N As Long')
[void]$m.AppendLine('End Type')
# procedures whose own code types every parameter
[void]$m.AppendLine("Public Sub U3(a As Long, b As Double, c As String)`r`n    gSink = a + b + Len(c)`r`nEnd Sub")
[void]$m.AppendLine("Public Sub U2(a As Long, ByVal k As Double, b As String)`r`n    gSink = a + k + Len(b)`r`nEnd Sub")
[void]$m.AppendLine("Public Function UF(a As Long, ByVal b As Double) As Double`r`n    UF = a * b`r`nEnd Function")
# small functions used inside arguments
[void]$m.AppendLine("Public Function FL(ByVal a As Long) As Long`r`n    FL = a + 1`r`nEnd Function")
[void]$m.AppendLine("Public Function FS(ByVal a As String) As String`r`n    FS = a & ""!""`r`nEnd Function")
[void]$m.AppendLine("Public Function FD(ByVal a As Double) As Double`r`n    FD = a / 2`r`nEnd Function")
[void]$m.AppendLine("Public Function K30f(ByVal a As Long, ByVal b As Long) As Long`r`n    K30f = a + b`r`nEnd Function")
foreach ($c in $sinks) {
    if ($c.N -eq 'K29') { continue }
    $end = if ($c.D -like 'Function*') { 'End Function' } elseif ($c.D -like 'Property*') { 'End Property' } else { 'End Sub' }
    # a Function returns something without touching its parameters
    $body = if ($c.D -like 'Function*As Long' -or $c.D -like 'Function*As Boolean') { "    $($c.N) = 1`r`n" } else { '' }
    [void]$m.AppendLine("Public $($c.D)`r`n$body$end")
}
foreach ($c in $passers) { [void]$m.AppendLine("Public $($c.D)`r`n    $($c.B)`r`nEnd Sub") }
[void]$m.AppendLine('Public Sub CallSinks(ByVal p As Long, ByVal q As Double, ByVal s As String, rp As Long, rs As String)')
[void]$m.AppendLine('    Dim n As Long, t As String, r As Rec, r2 As Rec, v As Variant, i As Long, arr() As Long')
[void]$m.AppendLine('    gSink = rp + Len(rs)')
[void]$m.AppendLine('    n = 11: t = "loc": r.D = 4.5: r.N = 6')
[void]$m.AppendLine('    ReDim arr(0 To 2): arr(0) = 10: arr(1) = 20: arr(2) = 30')
foreach ($c in $sinks) { [void]$m.AppendLine("    $($c.C)") }
[void]$m.AppendLine('End Sub')
[void]$m.AppendLine('Public Sub Drive()')
[void]$m.AppendLine('    Dim n2 As Long, t2 As String')
[void]$m.AppendLine('    n2 = 8: t2 = "rs"')
[void]$m.AppendLine('    CallSinks 3, 1.25, "sv", n2, t2')
foreach ($c in $passers) { [void]$m.AppendLine("    $($c.N) $($c.A)") }
[void]$m.AppendLine('End Sub')
# no VBA caller: Application.Run, and a cell
[void]$m.AppendLine("Public Sub RunSink(ByVal a As Long, ByVal b As String)`r`nEnd Sub")
[void]$m.AppendLine("Public Function CellSink(ByVal a As Long) As Long`r`n    CellSink = 1`r`nEnd Function")
[void]$m.AppendLine("Public Sub RunIt()`r`n    Application.Run ""RunSink"", 4, ""run""`r`nEnd Sub")
$moduleCode = $m.ToString()

$m2 = @'
Option Explicit
Public Sub X29(ByVal a As Long, ByVal b As Double, ByVal c As String)
End Sub
Public Sub U3M(a As Long, b As Double, c As String)
    M.gSink = a + b + Len(c)
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'Convoluted' @(@{ Kind=1; Name='M'; Code=$moduleCode }, @{ Kind=1; Name='M2'; Code=$m2 })
    $leaf = (Get-XRayMacroBook).Leaf
    $ws = $app.Workbooks.Item($leaf).Worksheets.Item(1)
    $ws.Range("B1").Value2 = 5
    # compiled first, as a saved workbook is: an uncompiled callee has no types to give
    $ctl = $app.VBE.CommandBars.FindControl(1, 578); if ($ctl -and $ctl.Enabled) { $ctl.Execute() }

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $app.Run($leaf + '!RunIt') | Out-Null
    $ws.Range("A1").Formula = "=CellSink(B1)"
    [void](Wait-XRayCalcDone $app)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $disarm = [string](Wait-LogLine $paths.Log 'VBA tracing: disarmed' $mark)

    $rows = @(Read-TraceRows $sx.ProcId)

    function Params([string]$fn) {
        $sig = @((SigOf $rows $fn) -split ',' | ForEach-Object { $_ -replace '\|.*$', '' })
        $vals = @([regex]::Matches([string](Remove-ArgAddress (ArgsOf $rows $fn)), 'a\d+:[^=]*=(.*?)(?= a\d+:|$)') | ForEach-Object { $_.Groups[1].Value })
        for ($i = 0; $i -lt $sig.Count; $i++) {
            [pscustomobject]@{ Type = ($sig[$i] -replace '#.*$', ''); Op = $(if ($sig[$i] -match '#(\d+)') { [int]$Matches[1] } else { 0 }); Value = $(if ($i -lt $vals.Count) { $vals[$i] } else { '' }) }
        }
    }
    function Expect([string]$fn, [object[]]$want) {
        $got = @(Params $fn)
        $bad = @()
        if ($got.Count -ne $want.Count) { $bad += "$($got.Count) parameter(s) read, $($want.Count) declared" }
        for ($i = 0; $i -lt $want.Count -and $i -lt $got.Count; $i++) {
            $g = $got[$i]; $w = [string]$want[$i]
            $typed = -not $g.Type.StartsWith('?')
            if ($w -eq '?') { if ($typed) { $bad += ("a{0}: typed {1}={2}, must stay untyped" -f ($i + 1), $g.Type, $g.Value) } ; continue }
            $wt, $wv = $w -split '=', 2
            if (-not $typed) { $bad += ("a{0}: untyped, want {1}" -f ($i + 1), $w); continue }
            if ($g.Type -cne $wt -or $g.Value -cne $wv) { $bad += ("a{0}: got {1}={2} want {3}" -f ($i + 1), $g.Type, $g.Value, $w) }
        }
        Write-Output ('  {0,-9} {1,-44} {2}' -f $fn, (SigOf $rows $fn), (Remove-ArgAddress (ArgsOf $rows $fn)))
        Check $fn ($bad.Count -eq 0) ("[$(SigOf $rows $fn)] " + $(if ($bad) { $bad -join ' | ' } else { 'as stated' }))
    }

    Write-Output ''
    foreach ($c in $sinks)   { if ($c.N -eq 'K29') { Expect 'X29' @('Long=3','Double=1.25','String="sv"') } else { Expect $c.N $c.W } }
    foreach ($c in $passers) { Expect $c.N $c.W }
    Expect 'RunSink' @('?','String="run"')
    Expect 'CellSink' @('?')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("{0} procedures, every parameter as stated" -f ($sinks.Count + $passers.Count + 2))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
