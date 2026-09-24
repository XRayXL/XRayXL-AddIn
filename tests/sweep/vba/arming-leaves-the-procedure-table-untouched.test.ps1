# Arming does not page in memory no procedure has used. The procedure table holds 16,384
# entries, about 11 MB, and clearing it by writing zeros would bring all of it into Excel's
# private working set at every arm. Checked on the first arm of a fresh session.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class XRayPrivateWs
{
    [StructLayout(LayoutKind.Sequential)]
    struct Counters
    {
        public uint cb, PageFaultCount;
        public UIntPtr PeakWorkingSetSize, WorkingSetSize, QuotaPeakPagedPoolUsage, QuotaPagedPoolUsage,
                       QuotaPeakNonPagedPoolUsage, QuotaNonPagedPoolUsage, PagefileUsage, PeakPagefileUsage,
                       PrivateUsage, PrivateWorkingSetSize, SharedCommitUsage;
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool K32GetProcessMemoryInfo(IntPtr process, ref Counters c, uint cb);
    public static long Bytes(IntPtr process)
    {
        var c = new Counters();
        c.cb = (uint)Marshal.SizeOf(typeof(Counters));
        if (!K32GetProcessMemoryInfo(process, ref c, c.cb)) throw new System.ComponentModel.Win32Exception();
        return (long)c.PrivateWorkingSetSize.ToUInt64();
    }
}
'@

$moduleCode = @'
Public Sub AT_Warm()
    Dim i As Long
    For i = 1 To 10
    Next i
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    # a later arm releases the previous table as it starts, which swamps the growth
    $earlierArms = @(Select-String -Path $paths.Log -SimpleMatch 'VBA tracing: ARMED' -ErrorAction SilentlyContinue).Count
    if ($earlierArms -gt 0) {
        Complete-Test -Skip -Detail "not the first VBA arm in this Excel ($earlierArms earlier), and only a first arm measures this"
    }

    New-XRayMacroBook $sx 'ArmMem' @(
        @{ Kind=1; Name='ArmMemCase'; Code=$moduleCode }
    )
    $leaf = (Get-XRayMacroBook).Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')
    [void](Set-XRayTraceParam $sx $null 'BUFFERSIZE' '0')

    # VBE7 loaded and the module compiled, so the arm is the only thing measured.
    $app.Run($leaf + '!AT_Warm') | Out-Null
    # held for the whole test: a handle off a collected Process object is closed
    $proc = [Diagnostics.Process]::GetProcessById($sx.ProcId)
    $handle = $proc.Handle

    $before = [XRayPrivateWs]::Bytes($handle)
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    $after = [XRayPrivateWs]::Bytes($handle)
    [GC]::KeepAlive($proc)
    $lossy = Stop-XRayTrace $sx
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $deltaMb = [math]::Round(($after - $before) / 1MB, 2)
    Write-XRayObservation 'private-working-set-growth-across-arm' "$deltaMb MB"
    Check 'arming-pages-in-less-than-4-mb' ($deltaMb -lt 4) `
          "private working set grew $deltaMb MB across the first arm (the whole table is ~11 MB)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "first arm grew the private working set by $deltaMb MB"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
