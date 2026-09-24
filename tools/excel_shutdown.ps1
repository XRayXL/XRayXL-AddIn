# Close the harness's Excel without a recovery prompt: a killed Excel keeps a copy of its
# open macro workbook, which the next Excel offers back behind a macro banner. With no pids
# given, it acts only on windowless Excels with our add-in loaded, sparing one in use.
[CmdletBinding()]
param(
    [int[]]$ProcessId = @(),
    [int]$TimeoutSeconds = 20,
    [switch]$Quiet
)
# the one pid-and-name kill and window-owner lookup, shared with StretchXL
$commonScript = Join-Path $PSScriptRoot '..\StretchXL\_common.ps1'
if (-not (Test-Path -LiteralPath $commonScript)) { throw "excel_shutdown: shared helpers not found at $commonScript" }
. $commonScript

$ExitPollMs = 400
# Excel writes its recovery files as it goes down, a moment after the process ends.
$RecoveryFileSettleMs = 800

function Write-Status($m) { if (-not $Quiet) { Write-Host $m } }

# Which Excels are ours to close?
$targets = @()
if ($ProcessId.Count) {
    foreach ($id in $ProcessId) {
        $p = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'EXCEL') { $targets += $p }
    }
} else {
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($p.MainWindowTitle) { continue }        # someone is looking at it
        $mine = $false
        try { $mine = @($p.Modules | Where-Object { $_.ModuleName -like '*XRayXL*' }).Count -gt 0 } catch {}
        if ($mine) { $targets += $p }
    }
}
if (-not $targets) { Write-Status "  no harness Excel to close"; return }

# GetActiveObject returns whichever Excel the ROT holds, which need not be ours,
# and closing its books discards unsaved work, so check its pid first.
try {
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    $xlPid = 0
    try { $xlPid = Get-WindowProcessId ([IntPtr]$xl.Hwnd) } catch {}

    if ($xlPid -le 0 -or -not ($targets.Id -contains $xlPid)) {
        Write-Status ("  the registered Excel is pid {0}, which is not one of ours -- leaving it alone" -f $xlPid)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    else {
        try { $xl.DisplayAlerts = $false } catch {}
        $closed = 0
        foreach ($wb in @($xl.Workbooks)) {
            # Saved = true stops the "do you want to save?" prompt; Close($false)
            # discards. A closed workbook cannot be recovered.
            try { $wb.Saved = $true; $wb.Close($false); $closed++ } catch {}
        }
        Write-Status ("  closed {0} workbook(s) before quitting pid {1}" -f $closed, $xlPid)
        try { $xl.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
} catch {
    Write-Status "  (no reachable COM instance; relying on Quit/wait per process)"
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt $deadline) {
    $alive = @($targets | ForEach-Object { Get-Process -Id $_.Id -ErrorAction SilentlyContinue })
    if (-not $alive) { break }
    Start-Sleep -Milliseconds $ExitPollMs
}
foreach ($p in $targets) {
    $still = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
    if (-not $still) { Write-Status ("  pid {0} quit cleanly" -f $p.Id); continue }
    if ($still.ProcessName -ne 'EXCEL') {
        Write-Status ("  pid {0} is now '{1}', not EXCEL -- the pid was recycled, leaving it alone" -f $p.Id, $still.ProcessName)
        continue
    }
    Write-Status ("  pid {0} did not quit in {1}s -- killing" -f $p.Id, $TimeoutSeconds)
    [void](Stop-ExcelByPid $p.Id)
}

# Leave nothing for Excel to offer the user next time. Matched by our pids
# (harness workbooks are <Name>_<pid>.xlsm): the folder also holds other people's work.
Start-Sleep -Milliseconds $RecoveryFileSettleMs
$ourMarks = @($targets | ForEach-Object { "*_$($_.Id)*" })
$rec = @(Get-ChildItem "$env:APPDATA\Microsoft\Excel" -File -ErrorAction SilentlyContinue |
         Where-Object { $f = $_; @($ourMarks | Where-Object { $f.Name -like $_ }).Count -gt 0 })
foreach ($f in $rec) {
    Write-Status ("  clearing recovery copy of our workbook: {0}" -f $f.Name)
    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
}
