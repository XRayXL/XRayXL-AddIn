# The whole-timeline assertion: one saved workbook with XLL and VBA UDFs
# produces ONE FILE, ONE MONOTONIC SEQUENCE, both sources, each call named and
# attributed to its calling cell. The Excel lifecycle, XLL registration,
# dialogs, deadlines and shutdown measurement all belong to StretchXL -- this
# drives and asserts, and nothing else.
#
# The MTC arm is a DIFFERENT test, not a repeat: single-threaded, write order
# and event order agree trivially; multithreaded they genuinely diverge, and
# what must still hold is everything that makes the file one trace.

function Invoke-TimelineTest {
    param([switch]$Mtc, [int]$ThreadSafeCells = 400, [int]$Threads = 8)

    . (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')

    # TxE is a test-XLL UDF taking and returning a double; VbaOuter calls it,
    # so one VBA activation contains an XLL one -- the nesting is the test.
    $VBA = @'
Public Function VbaOuter(ByVal n As Double) As Double
    VbaOuter = Application.Run("TxE", n) + 1
End Function

Public Function VbaPlain(ByVal n As Double) As Double
    VbaPlain = n * 3
End Function
'@

    try {
        $sx = Connect-TestExcel
        $app = $sx.App
        Set-XRaySessionDefaults $sx
        $app.EnableEvents = $false
        $paths = Get-XRayPaths $sx.ProcId

        if ($Mtc) {
            # Forced to a fixed thread count: "automatic" is free to decide one
            # thread is enough, which would make a green run mean nothing.
            $app.MultiThreadedCalculation.Enabled = $true
            $app.MultiThreadedCalculation.ThreadMode = 1
            $app.MultiThreadedCalculation.ThreadCount = $Threads
        }
        else {
            try { $app.MultiThreadedCalculation.Enabled = $false } catch {}
        }

        $formulas = @{
            'A1' = '=TxE(1.5)'       # XLL only
            'A2' = '=VbaPlain(2)'    # VBA only
            'A3' = '=VbaOuter(3)'    # VBA calling XLL: nesting
            'A4' = '=TxE(4.5)'
        }
        New-XRayMacroBook $sx 'OneTimeline' @(@{ Kind = 1; Name = 'Probe'; Code = $VBA }) $formulas {
            param($sheet)
            # TxThreadSafe is registered thread-safe ($), so Excel may spread this
            # block; without it, turning MTC on changes nothing.
            if ($Mtc) { $sheet.Range("C1:C$ThreadSafeCells").Formula = '=TxThreadSafe(ROW())' }
        } -SheetName 'Sheet1'
        $book = Get-XRayMacroBook
        $ws = $book.Sheet; $bookPath = $book.Path

        $mark = Get-LogLength $paths.Log
        $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
        if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
        # Awaited by its prefix, asserted by its content: every arm outcome begins
        # 'VBA tracing: ', and under the default VBA DEPTH=ALL only ARMED will do.
        $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
        if ($armLine -notmatch 'ARMED') {
            Complete-Test -Fail -Detail "did not arm: $armLine"
        }

        Invoke-XRayRecalc $app
        $mark2 = Get-LogLength $paths.Log
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }
        [void](Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2)

        $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)
        if ($rows.Count -eq 0) { Complete-Test -Fail -Detail "no trace rows at $(Get-XRayTraceCsv $sx.ProcId)" }

        $xllRows = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -or ($_.kind -eq 'exit' -and $_.source -eq 'XLL') })
        $vbaRows = @($rows | Where-Object { $_.source -eq 'VBA' })

        # ---- 1. one file, both sources ------------------------------------
        Check 'file-contains-xll-rows' ($xllRows.Count -gt 0) 'none found'
        Check 'file-contains-vba-rows' ($vbaRows.Count -gt 0) 'none found'

        # ---- 2. ONE MONOTONIC SEQUENCE ------------------------------------
        $seqs = @($rows | ForEach-Object { [int64]$_.seq })
        $dupes = @($seqs | Group-Object | Where-Object { $_.Count -gt 1 })
        Check 'no-seq-used-twice' ($dupes.Count -eq 0) `
              $(if ($dupes.Count) { "$($dupes.Count) duplicated, e.g. seq $($dupes[0].Name) x$($dupes[0].Count)" } else { '' })
        $mono = $true
        for ($i = 1; $i -lt $seqs.Count; $i++) { if ($seqs[$i] -le $seqs[$i - 1]) { $mono = $false; break } }
        Check 'seq-strictly-increasing' $mono 'it is not'

        # ---- 3. every entry paired with its exit --------------------------
        $bad = @()
        foreach ($g in ($rows | Where-Object { $_.span } | Group-Object span)) {
            $en = @($g.Group | Where-Object { $_.kind -eq 'entry' })
            $ex = @($g.Group | Where-Object { $_.kind -eq 'exit' })
            if ($en.Count -ne 1 -or $ex.Count -ne 1) { $bad += "span $($g.Name): $($en.Count)/$($ex.Count)"; continue }
            if ([int64]$ex[0].seq -le [int64]$en[0].seq) { $bad += "span $($g.Name): exit numbered before entry" }
        }
        Check 'every-span-entry-then-exit' ($bad.Count -eq 0) ($bad -join '; ')

        # ---- 4. THE NESTING, ACROSS SOURCES -------------------------------
        $outer = @($rows | Where-Object { $_.function -eq 'VbaOuter' })
        $oe = @($outer | Where-Object { $_.kind -eq 'entry' } | Select-Object -First 1)
        $ox = @($outer | Where-Object { $_.kind -eq 'exit' } | Select-Object -First 1)
        if (-not $oe -or -not $ox) {
            Check 'xll-nests-inside-vba' $false 'VbaOuter has no entry/exit pair'
        }
        else {
            $lo = [int64]$oe.seq; $hi = [int64]$ox.seq
            $inner = @($rows | Where-Object { $_.function -eq 'TxE' -and [int64]$_.seq -gt $lo -and [int64]$_.seq -lt $hi })
            Check 'xll-nests-inside-vba' ($inner.Count -ge 2) `
                  "found $($inner.Count) TxE row(s) between VbaOuter entry ($lo) and exit ($hi)"
        }

        # ---- 5. each named -------------------------------------------------
        $unnamed = @($rows | Where-Object { -not $_.function })
        Check 'every-row-names-its-function' ($unnamed.Count -eq 0) "$($unnamed.Count) unnamed"

        # ---- 6. each attributed to its calling cell -----------------------
        $entries = @($rows | Where-Object { $_.kind -eq 'entry' })
        $withCell = @($entries | Where-Object { Get-CallerCell $_ })
        Check 'cell-invoked-calls-carry-cell' ($withCell.Count -ge 4) `
              "only $($withCell.Count) of $($entries.Count) entries have one"
        $cellsSeen = @($withCell | ForEach-Object { Get-CallerCell $_ } | Sort-Object -Unique)
        Check 'vba-udf-attributed-to-A3' (@($cellsSeen | Where-Object { $_ -eq 'A3' }).Count -eq 1) 'A3 not among the cells seen'

        # ---- 6b. The caller invariants, and this scenario's value --------
        $cp = @(Test-RowInvariants $rows)
        Check 'caller-invariants-hold' ($cp.Count -eq 0) ($cp -join '; ')
        $offCell = @($entries | Where-Object { $_.caller -ne 'cell' })
        Check 'every-entry-called-by-its-cell' ($offCell.Count -eq 0) `
              "$($offCell.Count) entries off-cell, e.g. '$(@($offCell)[0].caller)'"

        # ---- MTC: did it engage, and did anything leak between threads? ---
        $threadIds = @($rows | ForEach-Object { $_.thread } | Sort-Object -Unique)
        if ($Mtc) {
            Check 'mtc-really-engaged' ($threadIds.Count -gt 1) `
                  "only $($threadIds.Count) thread(s) -- a run where MTC never engaged proves nothing"
            $split = @()
            foreach ($g in ($rows | Where-Object { $_.span } | Group-Object span)) {
                $t = @($g.Group | ForEach-Object { $_.thread } | Sort-Object -Unique)
                if ($t.Count -ne 1) { $split += "span $($g.Name) on $($t.Count) threads" }
            }
            Check 'no-span-split-across-threads' ($split.Count -eq 0) ($split -join '; ')
        }

        # ---- what makes sorting by qpc safe -------------------------------
        $backwards = @()
        foreach ($g in ($rows | Group-Object thread)) {
            $q = @($g.Group | ForEach-Object { [int64]$_.qpc })
            for ($i = 1; $i -lt $q.Count; $i++) {
                if ($q[$i] -lt $q[$i - 1]) { $backwards += "thread $($g.Name)"; break }
            }
        }
        Check 'qpc-monotonic-within-thread' ($backwards.Count -eq 0) ($backwards -join '; ')
        $unordered = @()
        foreach ($g in ($rows | Where-Object { $_.span } | Group-Object span)) {
            $en = @($g.Group | Where-Object { $_.kind -eq 'entry' } | Select-Object -First 1)
            $ex = @($g.Group | Where-Object { $_.kind -eq 'exit' } | Select-Object -First 1)
            if ($en -and $ex -and [int64]$ex.qpc -lt [int64]$en.qpc) { $unordered += "span $($g.Name)" }
        }
        Check 'span-stamped-entry-before-exit' ($unordered.Count -eq 0) ($unordered -join '; ')

        # seq vs qpc disagreement is REPORTED, not asserted: they answer
        # different questions and are allowed to differ (that is the log line
        # in the captured output, not a case).
        $inversions = 0
        for ($i = 1; $i -lt $rows.Count; $i++) {
            if ([int64]$rows[$i].qpc -lt [int64]$rows[$i - 1].qpc) { $inversions++ }
        }
        Write-Output ("note: qpc runs backwards at {0} of {1} row boundaries in file order (expected under MTC)" -f `
                      $inversions, ($rows.Count - 1))

        $checkFails = Get-XRayCheckFailures
        if ($checkFails -eq 0) {
            Complete-Test -Pass -Detail ("rows={0} threads={1} mtc={2}" -f $rows.Count, $threadIds.Count, [bool]$Mtc)
        }
        else {
            Complete-Test -Fail -Detail "$checkFails check(s) failed"
        }
    }
    catch {
        Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
    }
}
