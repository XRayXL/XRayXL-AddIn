# A parameter its own body never types -- unused, or only passed on to a typed ByRef parameter --
# takes the type its caller pushed: a literal's, a typed load's, or a local's, found at the call
# the caller is paused on. An entry VBA did not make itself, or an argument that is an expression,
# stays untyped. Diagnostics are on (suite.psd1), so a type the caller gave names the call, 1311.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

$moduleCode = @'
Public gN As Double, gS As String

' Nothing below types its parameters: they are unused, or only passed on.
Public Sub Lits(ByVal a As Long, ByVal s As String, ByVal d As Double, ByVal c As Currency, _
                ByVal g As Single, ByVal i As Integer, ByVal b As Boolean, ByVal l As LongLong)
End Sub
Public Sub Locals(x As Long, y As Double, z As String)
End Sub
Public Sub LocalV(v As Variant)
End Sub
Public Sub TakeD(y As Double)
    gN = y
End Sub
' Handed on to a typed ByRef parameter: no Variant, so no label.
Public Sub PassOn(x As Double)
    TakeD x
End Sub
Public Sub Chain1(k As Long)
    Chain2 k
End Sub
Public Sub Chain2(k As Long)
    Chain3 k
End Sub
Public Sub Chain3(k As Long)
    gN = k
End Sub

' Declined: an expression, a macro Excel ran, a UDF a cell called.
Public Sub Expr(ByVal a As Long)
End Sub
Public Sub RunTarget(ByVal k As Long)
End Sub
Public Function CellUnused(ByVal k As Long) As Long
    CellUnused = 1
End Function

Public Sub Drive()
    Dim n As Long, d As Double, s As String, vv As Variant
    n = 7: d = 2.5: s = "abc": vv = 5
    Lits 42, "hi", 2.5, 9.99@, 1.5!, 1000, True, 5^
    Locals n, d, s
    LocalV vv
    PassOn d
    Chain1 n
    Expr n + 1
    Application.Run "RunTarget", 5
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ByCaller' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    $ws = $app.Workbooks.Item($leaf).Worksheets.Item(1)

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $ws.Range("A1").Formula = "=CellUnused(4)"
    [void](Wait-XRayCalcDone $app)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $disarm = [string](Wait-LogLine $paths.Log 'VBA tracing: disarmed' $mark)

    $rows = @(Read-TraceRows $sx.ProcId)
    Write-Output ''
    foreach ($f in 'Lits','Locals','LocalV','PassOn','Chain1','Chain2','Chain3','Expr','RunTarget','CellUnused') {
        Write-Output ('  {0,-10} {1,-60} {2}' -f $f, (SigOf $rows $f), (ArgsOf $rows $f))
    }

    # Each parameter: its type, its opcode (1311 is the caller's call), and its value.
    function Params([string]$fn) {
        $sig = @((SigOf $rows $fn) -split ',' | ForEach-Object { $_ -replace '\|.*$', '' })
        $vals = @([regex]::Matches([string](ArgsOf $rows $fn), 'a\d+:[^=]*=(.*?)(?= a\d+:|$)') | ForEach-Object { $_.Groups[1].Value })
        for ($i = 0; $i -lt $sig.Count; $i++) {
            $t = $sig[$i] -replace '#.*$', ''
            $op = if ($sig[$i] -match '#(\d+)') { [int]$Matches[1] } else { 0 }
            [pscustomobject]@{ Type = $t; Op = $op; Value = $(if ($i -lt $vals.Count) { $vals[$i] } else { '' }) }
        }
    }
    # A third element names the opcode when the callee's own code types that parameter first.
    function Expect([string]$name, [string]$fn, [object[]]$want) {
        $got = @(Params $fn)
        $bad = @()
        for ($i = 0; $i -lt $want.Count; $i++) {
            $g = $got[$i]; $w = $want[$i]
            $op = if ($w.Count -gt 2) { $w[2] } else { 1311 }
            if (-not $g -or $g.Type -cne $w[0] -or $g.Op -ne $op -or $g.Value -cne $w[1]) {
                $bad += ("a{0}: got {1}#{2}={3} want {4}#{5}={6}" -f ($i + 1), $g.Type, $g.Op, $g.Value, $w[0], $op, $w[1])
            }
        }
        Check $name ($bad.Count -eq 0) ("$fn [$(SigOf $rows $fn)] " + $(if ($bad) { $bad -join ' | ' } else { 'all as planted' }))
    }

    # Literals, as the loads spell them: Boolean is Integer. A ByVal String is typed by its own
    # code even unused (663), so the caller is never asked.
    Expect 'literals-type-unused-byval-parameters' 'Lits' @(
        @('Long', '42'), @('String', '"hi"', 663), @('Double', '2.5'), @('Currency', '9.9900'),
        @('Single', '1.5'), @('Integer', '1000'), @('Integer', '-1'), @('LongLong', '5'))
    Expect 'typed-locals-type-unused-byref-parameters' 'Locals' @(@('Long&', '7'), @('Double&', '2.5'), @('String&', '"abc"'))
    Expect 'a-variant-local-types-an-unused-byref-variant' 'LocalV' @(,@('Variant&', 'Integer(5)'))
    Expect 'a-parameter-only-passed-on-to-a-typed-byref-is-typed' 'PassOn' @(,@('Double&', '2.5'))
    Expect 'a-chain-is-typed-at-every-level' 'Chain1' @(,@('Long&', '7'))
    Expect '...and-its-middle-level-from-one-level-up' 'Chain2' @(,@('Long&', '7'))

    foreach ($f in 'Expr', 'RunTarget', 'CellUnused') {
        $p = @(Params $f)[0]
        Check "not-typed-by-a-caller:$f" ($p -and $p.Op -ne 1311 -and $p.Type -like '?*') "$f [$(SigOf $rows $f)] $(ArgsOf $rows $f)"
    }

    $logged = if ($disarm -match '(\d+) parameter\(s\) typed by their caller') { [int]$Matches[1] } else { 0 }
    # Lits 7 (its String types itself), Locals 3, LocalV 1, PassOn 1, Chain1 1, Chain2 1
    Check 'the-disarm-line-counts-them' ($logged -eq 14) "disarm line: $logged"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("{0} parameters typed by their caller; Lits [{1}]" -f $logged, (ArgsOf $rows 'Lits'))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
