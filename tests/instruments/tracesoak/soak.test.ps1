# What one arming costs over a long session: arm once, then recalculate a complex sheet whose
# formulas reach XLL functions and VBA in another workbook. A per-call, per-row or per-span leak
# shows here.
#
# The same workload runs twice in the same process, disarmed and then armed, and the difference
# in growth per cycle is reported, because Excel grows on its own. Both phases follow a
# discarded warm-up. The control runs first and Excel's growth decelerates, so the attributed
# figure is conservative; a negative one means "not distinguishable from Excel".
#
# No threshold is asserted; it fails only if Excel dies or the workload cannot run. It is paced,
# because flat out it finds the limits of VBA's string space rather than anything about the
# tracer.
#
#    XRAY_SOAK_SECONDS   seconds per measured phase (default 180; two phases)
#    XRAY_SOAK_WARMUP    discarded warm-up seconds  (default 60)
#    XRAY_SOAK_CYCLE_MS  minimum ms between cycles  (default 200, i.e. 5/sec)

. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\sweep\_xray_common.ps1')
. (Join-Path $PSScriptRoot '..\_instrument.ps1')

$PhaseSeconds  = Get-EnvInt 'XRAY_SOAK_SECONDS'  180
$WarmupSeconds = Get-EnvInt 'XRAY_SOAK_WARMUP'    60
$CycleMs       = Get-EnvInt 'XRAY_SOAK_CYCLE_MS' 200
$SampleEvery   = 10          # cycles between resource samples

try {
    $sx  = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    Clear-XRayStaleTraces $sx.ProcId

    # ---- BOOK B: the OTHER workbook, holding VBA the model calls ----------
    # The cross-workbook case is the one a single-book soak would miss: the
    # calling cell lives in A, the procedure lives in B, and every entry row
    # has to attribute the frame to B while the caller stays A.
    $libCode = @'
Public Function LB_Price(ByVal r As Double, ByVal t As Double) As Double
    LB_Price = Exp(-r * t) * 100#
End Function
Public Function LB_Label(ByVal n As Double) As String
    LB_Label = "b-" & CStr(Int(n))
End Function
Public Function LB_Series(ByVal n As Long) As Variant
    Dim a() As Double, i As Long
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = Sqr(i) * 1.5
    Next i
    LB_Series = a
End Function
Public Function LB_Deep(ByVal n As Long) As Double
    If n > 0 Then LB_Deep = LB_Deep(n - 1) + 1 Else LB_Deep = 0
End Function
Public Function LB_Bag() As Variant
    Set LB_Bag = New Collection
End Function
'@
    New-XRayMacroBook $sx 'SoakLib' -SheetName 'L1' -Components @(
        @{ Kind = 1; Name = 'LM'; Code = $libCode }
    ) -Cells @{ 'A1' = '=LB_Price(0.03,2)'; 'A2' = '=LB_Deep(6)' }
    $libLeaf = (Get-XRayMacroBook 'SoakLib').Leaf

    # ---- BOOK A: the model -- XLL + own VBA + cross-book VBA --------------
    $modelCode = @'
Public Function MA_Vol() As Double
    Application.Volatile
    MA_Vol = Rnd()
End Function
Public Function MA_Leg(ByVal n As Long) As Double
    If n > 0 Then MA_Leg = MA_Leg(n - 1) + 1 Else MA_Leg = 0
End Function
Public Function MA_Chain(ByVal n As Long) As Double
    MA_Chain = MA_Leg(n) + MA_Leg(n \ 2)
End Function
Public Function MA_Text(ByVal n As Double) As String
    MA_Text = "m-" & CStr(Int(n)) & "-" & CStr(Int(Rnd() * 100))
End Function
Public Function MA_Basket() As Variant
    MA_Basket = Array(1.5, "two", CLng(3), True)
End Function
Public Function MA_Bag() As Variant
    Set MA_Bag = New Collection
End Function

' A MACRO, not a UDF: VBA the calculation engine never called. It is traced
' anyway because the interpreter hook is global, and it is half the workload
' a recalc-only soak would miss. It writes NOTHING to the sheet on purpose --
' a cell write would dirty the model and change how many recalcs a cycle
' costs, which would make the two phases' denominators different.
Public Sub MA_Work()
    Dim i As Long, s As String, c As Collection, d As Double
    Set c = New Collection
    For i = 1 To 40
        d = d + MA_Chain(6)
        s = MA_Text(CDbl(i))
        c.Add s
    Next i
    If c.Count <> 40 Then Err.Raise 5, , "soak macro did not run"
End Sub
'@
    # A calculation event: more VBA the engine reaches by a different route.
    $sheetCode = @'
Private Sub Worksheet_Calculate()
    Dim z As Double
    z = MA_Leg(5)
End Sub
'@
    # A COMPLEX SHEET: own VBA, cross-workbook VBA, and XLL functions in the
    # same dependency graph. TxCallsBack re-enters Excel through xlUDF, so
    # XLL frames genuinely nest while the VBA patches are live. Arrays spill,
    # so each gets its own column.
    $cells = @{
        'A1'  = '=MA_Vol()'
        'A2'  = '=MA_Chain(9)'
        'A3'  = '=MA_Text(7)'
        'A4'  = '=MA_Bag()'
        'A5'  = '=MA_Vol()'
        'A6'  = '=MA_Chain(4)'
        'A7'  = "='$libLeaf'!LB_Price(0.05,3)"
        'A8'  = "='$libLeaf'!LB_Label(42)"
        'A9'  = "='$libLeaf'!LB_Deep(8)"
        'A10' = "='$libLeaf'!LB_Bag()"
        'A11' = '=TxB(2,3)'
        'A12' = '=TxCallsBack(5)'
        'A13' = '=TxCallsBack2(5)'
        'C1'  = '=MA_Basket()'
        'F1'  = "='$libLeaf'!LB_Series(6)"
    }
    # Saved and reopened, so every asserted formula is as loaded.
    New-XRayMacroBook $sx 'SoakModel' -Cells $cells -Components @(
        @{ Kind = 1; Name = 'M'; Code = $modelCode }
        @{ Kind = 'Sheet'; Code = $sheetCode }
    )

    $threads = Enable-MultiThreadedCalc $app

    # One cycle = one full recalculation of both books plus one macro run, identical in both
    # phases. Rebuild, not Calculate: CalculateFull serves a non-volatile UDF from its last
    # result. Emits nothing, since it is called bare inside the phase loop and output would pile
    # up.
    function Invoke-SoakCycle {
        $t0 = [Diagnostics.Stopwatch]::StartNew()
        [void]$app.CalculateFullRebuild()
        [void]$app.Run('MA_Work')
        # Pace from the start of the cycle: a fixed sleep after it would make the phases differ
        # by the sleep as well as the work.
        $rest = $CycleMs - $t0.ElapsedMilliseconds
        if ($rest -gt 0) { Start-Sleep -Milliseconds $rest }
    }

    # Runs cycles for $Seconds and returns first/last samples plus the count. The first sample
    # is taken after the first $SampleEvery cycles, past the phase's own transient. Returns one
    # object and emits nothing; the caller prints progress.
    function Measure-SoakPhase([string]$Label, [int]$Seconds) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $n = 0; $first = $null; $last = $null
        $curve = [System.Collections.Generic.List[string]]::new()
        while ($sw.Elapsed.TotalSeconds -lt $Seconds) {
            Invoke-SoakCycle
            $n++
            if (($n % $SampleEvery) -eq 0) {
                $s = Get-ProcSample -ProcessId $sx.ProcId -At $n
                if (-not $s) { throw "excel is gone during the $Label phase at cycle $n" }
                if (-not $first) { $first = $s } else { $last = $s }
                # One line per sample, but only every 20th is kept: a 180s phase
                # at 200 cycles a second is thousands of samples, and a curve
                # nobody can read is just a bigger log.
                if ((($n / $SampleEvery) % 10) -eq 0) {
                    $curve.Add(("{0,-8} cycles={1,6} private={2}MB working={3}MB handles={4}" -f `
                                $Label, $n, $s.PrivateMB, $s.WorkingMB, $s.Handles))
                }
            }
        }
        $sw.Stop()
        if (-not $last) {
            throw ("the $Label phase completed only $n cycle(s) in {0:N0}s -- too few to measure a slope (need more than {1})" -f `
                   $sw.Elapsed.TotalSeconds, ($SampleEvery * 2))
        }
        [pscustomobject]@{
            Label = $Label; Cycles = $n; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            First = $first; Last = $last; Curve = $curve
            MBPer1000      = [math]::Round((($last.PrivateMB - $first.PrivateMB) / ($last.At - $first.At)) * 1000, 1)
            HandlesPer1000 = [math]::Round((($last.Handles   - $first.Handles)   / ($last.At - $first.At)) * 1000, 1)
        }
    }

    # ---- BOTH SOURCES ON, and CHECKED --------------------------------------
    # both sources must actually trace, or the figures measure nothing
    $bad = Set-XRayDepthAll $sx
    if ($bad) { Complete-Test -Fail -Detail $bad }

    # ---- PHASE 0: warm-up, DISCARDED --------------------------------------
    # First-touch costs -- JIT of the VBA, the XLL's first calls, Excel's own
    # caches -- all land here so neither measured phase carries them.
    [void](Measure-SoakPhase 'warmup' $WarmupSeconds)

    function Write-SoakCurve($Phase) {
        foreach ($line in $Phase.Curve) { Write-Output $line }
    }

    # ---- PHASE 1: the CONTROL. Loaded, not armed --------------------------
    # Loaded-but-disarmed is the right control: it isolates TRACING from
    # merely having the add-in in the process.
    $control = Measure-SoakPhase 'control' $PhaseSeconds
    Write-SoakCurve $control

    # ---- PHASE 2: ARMED, once, for the whole phase ------------------------
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm refused: $pressed" }
    $armed = Measure-SoakPhase 'armed' $PhaseSeconds
    Write-SoakCurve $armed

    # drops shrink the per-row denominator, and only the disarm reports them (-1 = disarm faulted)
    $dropped = Invoke-XRayDisarm $sx

    # ---- ROWS: the other denominator --------------------------------------
    # Counted once, after disarm, when the ring has drained. Per-cycle is what
    # the two phases share; per-row is what makes the figure comparable to a
    # differently shaped workload. Streamed, not loaded -- this file runs to
    # tens of megabytes.
    $rows = 0
    $csv = Get-XRayTraceCsv $sx.ProcId
    if (Test-Path $csv) {
        # Streamed in one pass: the file can run to gigabytes, and Get-Content
        # makes every line a PowerShell object first, which can exhaust the worker.
        $reader = [System.IO.File]::ReadLines($csv)
        foreach ($line in $reader) { $rows++ }
        $rows = [math]::Max(0, $rows - 1)   # the header is not a row
    }
    $mbPer100kRows = if ($rows -gt 0) {
        [math]::Round((($armed.Last.PrivateMB - $armed.First.PrivateMB) / $rows) * 100000, 2)
    } else { 0 }

    # ---- WHAT IS ATTRIBUTABLE TO TRACING ----------------------------------
    $mbAttr      = [math]::Round($armed.MBPer1000      - $control.MBPer1000, 1)
    $handlesAttr = [math]::Round($armed.HandlesPer1000 - $control.HandlesPer1000, 1)

    $dropNote = if ($dropped -lt 0) { ' DROPS=UNKNOWN (the disarm faulted inside the XLL)' }
                elseif ($dropped -gt 0) { " DROPPED={0} (per-row figure is a LOWER BOUND)" -f $dropped }
                else { '' }
    $detail = (("threads={0} phase_s={1} warmup_s={2} rows={3}$dropNote | " +
                "control: cycles={4} {5}MB/1000 {6}handles/1000 | " +
                "armed: cycles={7} {8}MB/1000 {9}handles/1000 | " +
                "ATTRIBUTABLE: {10}MB/1000 cycles, {11} handles/1000 cycles, {12}MB/100k rows") -f `
               $threads, $PhaseSeconds, $WarmupSeconds, $rows,
               $control.Cycles, $control.MBPer1000, $control.HandlesPer1000,
               $armed.Cycles, $armed.MBPer1000, $armed.HandlesPer1000,
               $mbAttr, $handlesAttr, $mbPer100kRows)

    # A soak that traced nothing measured nothing, whatever its numbers say.
    if ($rows -le 0) { Complete-Test -Fail -Detail ("armed phase produced NO trace rows -- " + $detail) }

    Complete-Test -Pass -Detail ("measured -- " + $detail)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' '))
}
