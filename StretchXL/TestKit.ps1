# StretchXL TestKit: a test dot-sources it to meet the contract in StretchXL.md.
#
#     . (Join-Path $PSScriptRoot '..\..\TestKit.ps1')   # or wherever it lives
#     $sx = Connect-TestExcel
#     ... do exactly one thing against $sx.App ...
#     Complete-Test -Pass            # or -Fail 'why' / -Skip 'why'
#
# With STRETCH_SESSION_HWND set, the kit binds to that Excel by window handle (never
# GetActiveObject, which binds "some Excel") and leaves the measured close to the
# manager. Without it, the kit starts its own Excel and closes it.

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# same suite config, window lookup and kill helpers as the manager
. (Join-Path $PSScriptRoot '_common.ps1')

# How long Connect-TestExcel keeps retrying the bind or the start: Excel can take
# several seconds to create its windows on a loaded machine.
$script:DefaultConnectTimeoutSeconds = 20
# How long a standalone close waits for Excel to exit before killing it: a clean
# quit takes a few seconds, and the verdict is already written.
$script:StandaloneQuitWaitSeconds = 30
# Poll intervals while waiting on Excel's windows, its pid, and its exit.
$script:BindPollMs = 200
$script:PidPollMs  = 100
$script:ExitPollMs = 250
# Settle after RegisterXLL when the suite does not say: add-in start-up is not instantaneous.
$script:DefaultSettleSeconds = 3

# --- native plumbing --------------------------------------------------------
# AccessibleObjectFromWindow(OBJID_NATIVEOM) on EXCEL7 gets one specific Excel's object
# model; EXCEL7 exists only once a workbook does, hence the baseline empty workbook.
if (-not ('SXKit.Native' -as [type])) {
    Add-Type -Name Native -Namespace SXKit -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern System.IntPtr FindWindowEx(System.IntPtr parent, System.IntPtr after, string cls, string title);
[System.Runtime.InteropServices.DllImport("oleacc.dll")]
public static extern int AccessibleObjectFromWindow(System.IntPtr hwnd, uint objectId, byte[] iid,
    [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)] out object ppv);
'@
}

$script:SxSession = $null

function Clear-ComStragglers {
    <#
    .SYNOPSIS
        Release un-rooted COM temporaries now. Each dot in $sx.App.Workbooks.Item(1)
        makes a wrapper nobody holds, and process exit does not reliably release
        them, leaving Excel alive after its close. Variables in scope stay the
        test's to release.
    #>
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}

function Get-TestWorkDir {
    <#
      One directory per test, inside the session's add-in output directory.
      No fallback between managed and standalone: a missing value under the
      manager means the manager is broken, and writing elsewhere would hide it.
    #>
    param([Parameter(Mandatory)][bool]$Managed)

    if ($Managed) {
        $root = $env:STRETCH_SESSION_DIR
        if ([string]::IsNullOrEmpty($root)) {
            throw "STRETCH_SESSION_DIR is not set. The manager sets it before every test; running under a manager without it means the session was never opened."
        }
        $id = $env:STRETCH_TEST_ID
        if ([string]::IsNullOrEmpty($id)) {
            throw "STRETCH_TEST_ID is not set. The manager sets it from the test's filename before the spawn; without it this test cannot name its own work directory."
        }
    }
    else {
        $root = Join-Path $env:TEMP 'StretchXL\standalone-work'
        $self = (Get-PSCallStack | Where-Object { $_.ScriptName -like '*.test.ps1' } |
                 Select-Object -First 1).ScriptName
        if ([string]::IsNullOrEmpty($self)) {
            throw "Cannot identify the running test: no *.test.ps1 on the call stack. Connect-TestExcel must be called from a .test.ps1 file."
        }
        $id = [IO.Path]::GetFileNameWithoutExtension($self) -replace '\.test$',''
    }

    $dir = Join-Path $root $id
    New-Item -ItemType Directory -Force $dir | Out-Null
    return $dir
}

function Connect-TestExcel {
    <#
    .SYNOPSIS
        Get the Excel this test should use: the manager's session when one is
        designated in the environment, otherwise a standalone Excel the kit
        starts (and will close) itself. Returns a session object:
        .App (Application), .ProcId, .Hwnd, .Managed, .WorkDir.
    #>
    [CmdletBinding()]
    param(
        # How long to keep retrying the bind or the start before failing.
        [int]$ConnectTimeoutSeconds = $script:DefaultConnectTimeoutSeconds
    )

    $sessionHwnd = $env:STRETCH_SESSION_HWND
    $deadline = (Get-Date).AddSeconds($ConnectTimeoutSeconds)

    if (-not [string]::IsNullOrEmpty($sessionHwnd)) {
        # ---- managed: bind to exactly the designated instance -------------
        $mainHwnd = [IntPtr][int64]$sessionHwnd
        $wantPid  = [int]$env:STRETCH_SESSION_PID

        # A stale handle reused by another window must fail here, not bind
        # this test to somebody else's Excel.
        $gotPid = Get-WindowProcessId $mainHwnd
        if ($gotPid -ne $wantPid) {
            throw "STRETCH_SESSION_HWND $sessionHwnd belongs to pid $gotPid, not the designated pid $wantPid -- refusing to bind"
        }

        # Two PowerShell 5.1 traps: $null for a P/Invoke string marshals as ""
        # ([NullString]::Value is the real null), and 0xFFFFFFF0 parses as int32 -16.
        $OBJID_NATIVEOM = [uint32]4294967280   # 0xFFFFFFF0
        $windowObj = $null
        $iidDispatch = ([Guid]'00020400-0000-0000-C000-000000000046').ToByteArray()
        while ((Get-Date) -lt $deadline -and $null -eq $windowObj) {
            $xlDesk = [SXKit.Native]::FindWindowEx($mainHwnd, [IntPtr]::Zero, 'XLDESK', [NullString]::Value)
            if ($xlDesk -ne [IntPtr]::Zero) {
                $excel7 = [SXKit.Native]::FindWindowEx($xlDesk, [IntPtr]::Zero, 'EXCEL7', [NullString]::Value)
                if ($excel7 -ne [IntPtr]::Zero) {
                    $raw = $null
                    $hr = [SXKit.Native]::AccessibleObjectFromWindow($excel7, $OBJID_NATIVEOM, $iidDispatch, [ref]$raw)
                    if ($hr -eq 0 -and $null -ne $raw) { $windowObj = $raw }
                }
            }
            if ($null -eq $windowObj) { Start-Sleep -Milliseconds $script:BindPollMs }
        }
        if ($null -eq $windowObj) {
            throw "could not bind to the designated Excel (pid $wantPid) within ${ConnectTimeoutSeconds}s -- no EXCEL7 window / OBJID_NATIVEOM refused"
        }

        $appObj = $windowObj.Application

        # The manager's baseline book carries the bind window, so it is the one
        # book Complete-Test must not close. Best-effort: unknown means the
        # narrower "has a Path" rule applies instead.
        $baseBook = $null
        try { $baseBook = [string]$windowObj.Parent.Name } catch {}

        $script:SxSession = @{
            App = $appObj; Window = $windowObj; ProcId = $wantPid; Hwnd = $mainHwnd
            Managed = $true; WorkDir = (Get-TestWorkDir -Managed $true); BaseBook = $baseBook
        }
    }
    else {
        # ---- standalone: start our own, ledger it, and own its close ------
        $ledger = $env:STRETCH_LEDGER
        $ownLedger = $false
        if ([string]::IsNullOrEmpty($ledger)) {
            $ledgerHome = Join-Path $env:TEMP 'StretchXL'
            New-Item -ItemType Directory -Force $ledgerHome | Out-Null
            # named like the manager's ledgers, so StretchXL.ps1 -Cleanup sweeps it too
            $ledger = Join-Path $ledgerHome "StretchXL_pids.standalone.$PID.txt"
            # ours to delete once its Excel is gone
            $ownLedger = $true
        }

        # Match the manager: config and SessionEnvironment before Excel starts,
        # because a COM-created Excel inherits its creator's environment.
        $testScript = $null
        foreach ($fr in (Get-PSCallStack)) {
            if ($fr.ScriptName -and $fr.ScriptName -like '*.test.ps1') { $testScript = $fr.ScriptName }
        }
        $workDir = Get-TestWorkDir -Managed $false
        $suiteCfg = @{}
        if ($testScript) {
            $ownDir = Split-Path $testScript -Parent
            # no tests root standalone; the resolver still climbs the suite.psd1 chain
            $suiteCfg = Resolve-SuiteConfig -SuiteDir $ownDir -RootDir $ownDir
            Assert-SuiteArtifacts -SuiteDir $ownDir -Config $suiteCfg
        }
        # one test per session standalone, so the work dir is the session dir
        $env:STRETCH_SESSION_DIR = $workDir
        if ($suiteCfg['SessionEnvironment']) {
            foreach ($name in @($suiteCfg['SessionEnvironment'].Keys)) {
                $val = ([string]$suiteCfg['SessionEnvironment'][$name]).Replace('{SessionDir}', $workDir)
                [Environment]::SetEnvironmentVariable($name, $val, 'Process')
            }
        }

        $appObj = New-Object -ComObject Excel.Application
        $mainHwnd = [IntPtr]::Zero; $ownPid = 0
        while ((Get-Date) -lt $deadline -and $ownPid -eq 0) {
            $rawHwnd = [int64]0
            try { $rawHwnd = [int64]$appObj.Hwnd } catch {}
            if ($rawHwnd -ne 0) {
                $found = Get-WindowProcessId ([IntPtr]$rawHwnd)
                if ($found -ne 0) { $mainHwnd = [IntPtr]$rawHwnd; $ownPid = $found }
            }
            if ($ownPid -eq 0) { Start-Sleep -Milliseconds $script:PidPollMs }
        }
        if ($ownPid -eq 0) { throw 'standalone Excel started but its pid could not be learned -- an untracked Excel may remain' }
        Add-Content -Path $ledger -Value $ownPid -Encoding ascii

        # Same baseline as a managed session: one empty workbook.
        $booksObj = $appObj.Workbooks
        [void]$booksObj.Add()

        # existence was checked above, before Excel started
        $xllList = @($suiteCfg['RegisterXll'] | Where-Object { $_ })
        $settle  = if ($suiteCfg.ContainsKey('RegisterXllSettleSeconds')) { [int]$suiteCfg['RegisterXllSettleSeconds'] } else { $script:DefaultSettleSeconds }
        foreach ($xllPath in $xllList) {
            if (-not [bool]$appObj.RegisterXLL($xllPath)) { throw "RegisterXLL failed: $xllPath" }
        }
        if ($xllList.Count -gt 0) { Start-Sleep -Seconds $settle }

        $script:SxSession = @{
            App = $appObj; Window = $null; Books = $booksObj; ProcId = $ownPid; Hwnd = $mainHwnd
            Managed = $false; WorkDir = $workDir; Ledger = $ledger; OwnLedger = $ownLedger
        }
    }
    return $script:SxSession
}

function Get-SessionDialogs {
    <#
    .SYNOPSIS
        The dialogs the manager's watchdog has dismissed so far on this worker,
        newest last. A test that expected them calls Write-DialogsHandled to
        keep its PASS; otherwise the manager turns it into a FAIL.
    #>
    if ([string]::IsNullOrEmpty($env:STRETCH_DIALOG_LOG)) { return @() }
    return @(Get-Content $env:STRETCH_DIALOG_LOG -ErrorAction SilentlyContinue)
}

function Write-DialogsHandled {
    <#
    .SYNOPSIS
        Declare that this test saw and judged the session's dialogs itself.
        The dialog count still lands on the record; only the automatic
        PASS-to-FAIL conversion is waived.
    #>
    Write-Output 'STRETCH dialogs=handled'
}

# The manager keeps the last 8192 bytes of a test's output, so a detail longer than
# that cuts off the start of its own verdict line and the result reads as ERROR.
$script:DetailMaxChars = 2000
function Limit-Detail([string]$Text) {
    if ($Text.Length -le $script:DetailMaxChars) { return $Text }
    return $Text.Substring(0, $script:DetailMaxChars) + " ...[cut $($Text.Length - $script:DetailMaxChars) chars]"
}

function Write-TestCase {
    <#
    .SYNOPSIS
        Report one named case inside a multi-case test. Emit before
        Complete-Test; the manager counts them per result.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$Pass,
        [switch]$Fail,
        [string]$Detail = ''
    )
    if ($Pass -eq $Fail) { throw 'Write-TestCase: exactly one of -Pass / -Fail' }
    $cv = if ($Pass) { 'PASS' } else { 'FAIL' }
    $cleanName = ($Name -replace '\s+', '_')
    $cleanDetail = Limit-Detail (($Detail -replace '\s+', ' ').Trim())
    $caseLine = "STRETCH case=$cleanName verdict=$cv"
    if ($cleanDetail) { $caseLine += " detail=$cleanDetail" }
    Write-Output $caseLine
}

function Complete-Test {
    <#
    .SYNOPSIS
        Emit the verdict line, release the kit's COM references, close a
        standalone Excel, and exit 0: the verdict line carries the result, and
        a nonzero exit means the test crashed.
    #>
    [CmdletBinding()]
    param(
        [switch]$Pass,
        [switch]$Fail,
        [switch]$Skip,
        [string]$Detail = ''
    )
    $picked = @($Pass, $Fail, $Skip | Where-Object { $_ }).Count
    if ($picked -ne 1) { throw 'Complete-Test: exactly one of -Pass / -Fail / -Skip' }
    $tv = if ($Pass) { 'PASS' } elseif ($Fail) { 'FAIL' } else { 'SKIP' }

    $cleanDetail = Limit-Detail (($Detail -replace '\s+', ' ').Trim())
    $verdictLine = "STRETCH verdict=$tv"
    if ($cleanDetail) { $verdictLine += " detail=$cleanDetail" }
    Write-Output $verdictLine

    $sess = $script:SxSession
    if ($null -ne $sess) {
        if (-not $sess.Managed) {
            # Quit, release, bounded wait, and kill (by pid and name) only what
            # we started. Best-effort: the verdict is already on stdout.
            try { $sess.App.Quit() } catch {}
            foreach ($comRef in @($sess.Books, $sess.App)) {
                if ($comRef) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comRef) } catch {} }
            }
            $script:SxSession = $null
            Clear-ComStragglers
            $waitUntil = (Get-Date).AddSeconds($script:StandaloneQuitWaitSeconds)
            while ((Get-Date) -lt $waitUntil -and (Get-Process -Id $sess.ProcId -ErrorAction SilentlyContinue)) {
                Start-Sleep -Milliseconds $script:ExitPollMs
            }
            [void](Stop-ExcelByPid $sess.ProcId)
            # a ledger the manager handed us is the manager's to clear
            if ($sess.OwnLedger -and $sess.Ledger) {
                Remove-Item -LiteralPath $sess.Ledger -Force -ErrorAction SilentlyContinue
            }
        }
        else {
            # A book left open is recalculated into the next test's trace, so every book but the
            # baseline goes, matched by name: an unsaved Workbooks.Add() looks just like it.
            # Without a known baseline, only books with a Path are closed.
            try {
                $ap = $sess.App
                if ($ap) {
                    $savedAlerts = $true
                    try { $savedAlerts = $ap.DisplayAlerts } catch {}
                    try { $ap.DisplayAlerts = $false } catch {}
                    $wbs = $ap.Workbooks
                    $wi = 1
                    while ($wi -le $wbs.Count) {
                        $wbx = $wbs.Item($wi)
                        $hasPath = $false
                        try { $hasPath = -not [string]::IsNullOrEmpty([string]$wbx.Path) } catch {}

                        $isBaseline = $false
                        if ($sess.BaseBook) {
                            try { $isBaseline = ([string]$wbx.Name -eq [string]$sess.BaseBook) } catch {}
                        }
                        $closeIt = if ($sess.BaseBook) { -not $isBaseline } else { $hasPath }

                        if ($closeIt) {
                            try { $wbx.Saved = $true } catch {}
                            try { $wbx.Close($false) } catch {}
                            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbx)
                            # collection shrank -- do not advance the index
                        }
                        else {
                            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbx)
                            $wi++
                        }
                    }
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbs)
                    try { $ap.DisplayAlerts = $savedAlerts } catch {}
                }
            } catch {}

            # Release only what the kit acquired, then sweep. The manager
            # closes and measures the process after this one exits.
            foreach ($comRef in @($sess.App, $sess.Window)) {
                if ($comRef) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comRef) } catch {} }
            }
            $script:SxSession = $null
            Clear-ComStragglers
        }
    }
    exit 0
}
