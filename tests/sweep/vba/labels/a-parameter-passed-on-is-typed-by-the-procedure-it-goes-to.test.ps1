# A parameter the body only passes on by address, to a VBA procedure its own pool names, takes
# that procedure's type for the slot: a ByRef argument must match its parameter exactly. It needs
# no caller, so it types a UDF a cell called, a macro Application.Run ran, and a procedure called
# with expressions. Diagnostics are on (suite.psd1): such a type names the push, 751 or 671.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

$moduleCode = @'
Public gN As Double

Public Function Helper(y As Double) As Double
    Helper = y * 2
End Function
Public Sub TakeL(y As Long)
    gN = y
End Sub
Public Sub TakeS(y As String)
    gN = Len(y)
End Sub

' A cell calls it, so there is no VBA caller: only where x goes can type it.
Public Function CellPass(x As Double) As Double
    CellPass = Helper(x)
End Function
' Application.Run runs it: its ByVal slot's address goes to a ByRef Long.
Public Sub RunPass(ByVal k As Long)
    TakeL k
End Sub
' Its caller passes expressions, which that call does not type.
Public Sub ExprPass(x As Double, s As String)
    gN = Helper(x)
    TakeS s
End Sub
' Passed to a procedure that only passes it on in turn: one level is read, so it stays untyped.
Public Sub PassToPasser(x As Long)
    Passer x
End Sub
Public Sub Passer(y As Long)
    TakeL y
End Sub

Public Sub Drive()
    Dim d As Double, s As String, n As Long
    d = 2.5: s = "abc": n = 3
    ExprPass d + 0, s & ""
    Application.Run "RunPass", 5
    PassToPasser n + 0
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ByCallee' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    $ws = $app.Workbooks.Item($leaf).Worksheets.Item(1)
    $ws.Range("B1").Value2 = 4
    # Compiled first, as a saved workbook is: until then a pool entry names the callee's
    # compile-on-demand stub, whose types do not exist yet.
    $ctl = $app.VBE.CommandBars.FindControl(1, 578); if ($ctl -and $ctl.Enabled) { $ctl.Execute() }

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $ws.Range("A1").Formula = "=CellPass(B1)"
    [void](Wait-XRayCalcDone $app)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $disarm = [string](Wait-LogLine $paths.Log 'VBA tracing: disarmed' $mark)

    $rows = @(Read-TraceRows $sx.ProcId)
    Write-Output ''
    foreach ($f in 'CellPass','RunPass','ExprPass','PassToPasser','Passer') {
        Write-Output ('  {0,-12} {1,-40} {2}' -f $f, (SigOf $rows $f), (Remove-ArgAddress (ArgsOf $rows $f)))
    }

    function Params([string]$fn) {
        $sig = @((SigOf $rows $fn) -split ',' | ForEach-Object { $_ -replace '\|.*$', '' })
        $vals = @([regex]::Matches([string](Remove-ArgAddress (ArgsOf $rows $fn)), 'a\d+:[^=]*=(.*?)(?= a\d+:|$)') | ForEach-Object { $_.Groups[1].Value })
        for ($i = 0; $i -lt $sig.Count; $i++) {
            [pscustomobject]@{ Type = ($sig[$i] -replace '#.*$', ''); Op = $(if ($sig[$i] -match '#(\d+)') { [int]$Matches[1] } else { 0 }); Value = $(if ($i -lt $vals.Count) { $vals[$i] } else { '' }) }
        }
    }
    function Expect([string]$name, [string]$fn, [object[]]$want) {
        $got = @(Params $fn)
        $bad = @()
        for ($i = 0; $i -lt $want.Count; $i++) {
            $g = $got[$i]; $w = $want[$i]
            if (-not $g -or $g.Type -cne $w[0] -or $g.Op -ne $w[1] -or $g.Value -cne $w[2]) {
                $bad += ("a{0}: got {1}#{2}={3} want {4}#{5}={6}" -f ($i + 1), $g.Type, $g.Op, $g.Value, $w[0], $w[1], $w[2])
            }
        }
        Check $name ($bad.Count -eq 0) ("$fn [$(SigOf $rows $fn)] " + $(if ($bad) { $bad -join ' | ' } else { 'all as declared' }))
    }

    Expect 'a-udf-a-cell-called-is-typed-by-what-it-calls' 'CellPass' @(,@('Double&', 751, '4'))
    Expect 'a-macro-application-run-ran-is-typed-by-what-it-calls' 'RunPass' @(,@('Long', 671, '5'))
    Expect 'expression-arguments-are-typed-by-where-they-go' 'ExprPass' @(@('Double&', 751, '2.5'), @('String&', 751, '"abc"'))
    Expect 'one-level-down-types-the-passer' 'Passer' @(,@('Long&', 751, '3'))
    $p = @(Params 'PassToPasser')[0]
    Check 'two-levels-down-is-not-read' ($p -and $p.Type -like '?*') "PassToPasser [$(SigOf $rows 'PassToPasser')]"

    $logged = if ($disarm -match '(\d+) parameter\(s\) typed by the procedure they are passed to') { [int]$Matches[1] } else { 0 }
    # CellPass 1, RunPass 1, ExprPass 2, Passer 1
    Check 'the-disarm-line-counts-them' ($logged -eq 5) "disarm line: $logged"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("{0} parameters typed by where they go; CellPass [{1}]" -f $logged, (ArgsOf $rows 'CellPass'))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
