# What one arming costs over a long session, where a per-call, per-row or per-span leak shows. The
# workload runs disarmed then armed after a discarded warm-up, and growth per cycle is compared
# because Excel grows on its own; nothing is asserted. Paced, or it finds VBA's string-space limits.
#
#    XRAY_SOAK_SECONDS   seconds per measured phase (default 180; two phases)
#    XRAY_SOAK_WARMUP    discarded warm-up seconds  (default 60)
#    XRAY_SOAK_CYCLE_MS  minimum ms between cycles  (default 200, i.e. 5/sec)
#    XRAY_SOAK_EVENTS    All (default), Calc, Selection, CalcAndSelection or None: Excel's events
#                        recorded while armed; every cycle raises ten kinds either way

. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\sweep\_xray_common.ps1')
. (Join-Path $PSScriptRoot '..\_instrument.ps1')

$PhaseSeconds  = Get-EnvInt 'XRAY_SOAK_SECONDS'  180
$WarmupSeconds = Get-EnvInt 'XRAY_SOAK_WARMUP'    60
$CycleMs       = Get-EnvInt 'XRAY_SOAK_CYCLE_MS' 200
$SampleEvery   = 10          # cycles between resource samples
$EventSet      = if ($env:XRAY_SOAK_EVENTS) { $env:XRAY_SOAK_EVENTS } else { 'All' }

try {
    $sx  = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    Clear-XRayStaleTraces $sx.ProcId

    # ---- book B: the other workbook, holding VBA the model calls ----------
    # A single-book soak would miss this: the calling cell lives in A, the procedure in B.
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

    # ---- book A: the model -- XLL + own VBA + cross-book VBA --------------
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
    # Own VBA, cross-workbook VBA and XLL functions in one dependency graph. TxCallsBack re-enters
    # Excel through xlUDF, so XLL frames nest while the VBA patches are live. Arrays spill, so each
    # gets its own column.
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

    # A book of its own to raise events in, so the model's formulas never read what it writes.
    # Its edits recalculate the volatile cells too, in both phases alike.
    $scratch = New-EventScratchBook $sx
    $model = (Get-XRayMacroBook 'SoakModel').Book
    $modelLeaf = (Get-XRayMacroBook 'SoakModel').Leaf
    $model.Activate()
    $script:round = 0

    $threads = Enable-MultiThreadedCalc $app

    # One cycle is a full rebuild of both books plus one macro run. Rebuild, not Calculate:
    # CalculateFull serves a non-volatile UDF from its last result. Emits nothing, or output piles up.
    function Invoke-SoakCycle {
        $t0 = [Diagnostics.Stopwatch]::StartNew()
        [void]$app.CalculateFullRebuild()
        [void]$app.Run("'$modelLeaf'!MA_Work")      # named: with three books open, a bare name may miss
        $script:round++
        Invoke-EventPoke $sx $scratch $model $script:round
        # Pace from the start of the cycle: a fixed sleep after it would make the phases differ
        # by the sleep as well as the work.
        $rest = $CycleMs - $t0.ElapsedMilliseconds
        if ($rest -gt 0) { Start-Sleep -Milliseconds $rest }
    }

    # The first sample follows the first $SampleEvery cycles, past the phase's own transient.
    # Returns one object and emits nothing; the caller prints progress.
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
                # Only every tenth sample goes into the curve: a curve nobody can read is just a bigger log.
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

    # ---- both sources on, and checked --------------------------------------
    # both sources must actually trace, or the figures measure nothing
    $bad = Set-XRayDepthAll $sx
    if ($bad) { Complete-Test -Fail -Detail $bad }
    $chosen = switch ($EventSet) {
        'All'              { $script:AppEvents }
        'Calc'             { $script:CalcEvents }
        'Selection'        { $script:SelectionEvents }
        'CalcAndSelection' { $script:CalcEvents + $script:SelectionEvents }
        'None'             { @() }
        default            { Complete-Test -Fail -Detail "XRAY_SOAK_EVENTS must be All, Calc, Selection, CalcAndSelection or None, not '$EventSet'" }
    }
    $bad = @(Set-EventChoice $sx $chosen @{})
    if ($bad.Count) { Complete-Test -Fail -Detail ("events setter: " + ($bad -join ' | ')) }

    # ---- phase 0: warm-up, discarded --------------------------------------
    # First-touch costs -- JIT of the VBA, the XLL's first calls, Excel's own
    # caches -- all land here so neither measured phase carries them.
    [void](Measure-SoakPhase 'warmup' $WarmupSeconds)

    function Write-SoakCurve($Phase) {
        foreach ($line in $Phase.Curve) { Write-Output $line }
    }

    # ---- phase 1: the control, loaded but not armed -----------------------
    # This isolates tracing from merely having the add-in in the process.
    $control = Measure-SoakPhase 'control' $PhaseSeconds
    Write-SoakCurve $control

    # ---- phase 2: armed, once, for the whole phase ------------------------
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm refused: $pressed" }
    $armed = Measure-SoakPhase 'armed' $PhaseSeconds
    Write-SoakCurve $armed

    # drops shrink the per-row denominator, and only the disarm reports them (-1 = disarm faulted)
    $dropped = Invoke-XRayDisarm $sx

    # ---- rows: the other denominator --------------------------------------
    # Per-row makes the figure comparable to a differently shaped workload. Counted after disarm,
    # once the ring has drained.
    $rows = 0
    $csv = Get-XRayTraceCsv $sx.ProcId
    if (Test-Path $csv) {
        # Streamed in one pass: the file can run to gigabytes, and Get-Content
        # makes every line a PowerShell object first, which can exhaust the worker.
        $reader = [System.IO.File]::ReadLines($csv)
        foreach ($line in $reader) { $rows++ }
        $rows = [math]::Max(0, $rows - 1)   # the header is not a row
    }
    # Every chosen kind the cycle raises must be in the trace, or the soak did not exercise it.
    $events = Get-EventCounts $csv
    $eventRows = 0; foreach ($v in $events.Values) { $eventRows += $v }
    $missing = @($script:PokedEvents | Where-Object { $chosen -contains $_ -and -not $events[$_] })
    $stray = @($events.Keys | Where-Object { $chosen -notcontains $_ })
    $eventNote = "events=$EventSet rows=$eventRows (" +
                 (($events.Keys | Sort-Object | ForEach-Object { "$_=$($events[$_])" }) -join ' ') + ")"

    $mbPer100kRows = if ($rows -gt 0) {
        [math]::Round((($armed.Last.PrivateMB - $armed.First.PrivateMB) / $rows) * 100000, 2)
    } else { 0 }

    # ---- what is attributable to tracing ----------------------------------
    $mbAttr      = [math]::Round($armed.MBPer1000      - $control.MBPer1000, 1)
    $handlesAttr = [math]::Round($armed.HandlesPer1000 - $control.HandlesPer1000, 1)

    $dropNote = if ($dropped -lt 0) { ' DROPS=UNKNOWN (the disarm faulted inside the XLL)' }
                elseif ($dropped -gt 0) { " DROPPED={0} (per-row figure is a LOWER BOUND)" -f $dropped }
                else { '' }
    $detail = (("threads={0} phase_s={1} warmup_s={2} rows={3}$dropNote | $eventNote | " +
                "control: cycles={4} {5}MB/1000 {6}handles/1000 | " +
                "armed: cycles={7} {8}MB/1000 {9}handles/1000 | " +
                "ATTRIBUTABLE: {10}MB/1000 cycles, {11} handles/1000 cycles, {12}MB/100k rows") -f `
               $threads, $PhaseSeconds, $WarmupSeconds, $rows,
               $control.Cycles, $control.MBPer1000, $control.HandlesPer1000,
               $armed.Cycles, $armed.MBPer1000, $armed.HandlesPer1000,
               $mbAttr, $handlesAttr, $mbPer100kRows)

    # A soak that traced nothing measured nothing, whatever its numbers say.
    if ($rows -le 0) { Complete-Test -Fail -Detail ("armed phase produced NO trace rows -- " + $detail) }
    if ($missing.Count) { Complete-Test -Fail -Detail ("chosen events raised every cycle are missing: $($missing -join ',') -- " + $detail) }
    if ($stray.Count) { Complete-Test -Fail -Detail ("events recorded though not chosen: $($stray -join ',') -- " + $detail) }

    Complete-Test -Pass -Detail ("measured -- " + $detail)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' '))
}
