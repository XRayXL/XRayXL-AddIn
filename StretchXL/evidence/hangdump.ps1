# Dump a process that will not exit, while it is still stuck: every thread is
# parked where it hung, and a killed process leaves no evidence at all.
#
# Dumps by pid only and never kills anything; the caller decides that.
# Writes one object to the output stream: ProcessId, Path ($null when no dump
# was written), SizeMB and Error. Nothing goes to any other stream, so a job
# or a pipeline can use the result directly.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][int]$ProcessId,
    [string]$Path,
    # Full memory keeps evidence/dumpstack.py working (it reads the
    # Memory64List stream). -Small keeps stacks and handles: about 1/20th the
    # size, enough to see where a thread is stuck but not what it holds.
    [switch]$Small
)
$ErrorActionPreference = "Stop"

function New-DumpResult([string]$DumpPath, $SizeMB, [string]$Why) {
    [pscustomobject]@{ ProcessId = $ProcessId; Path = $DumpPath; SizeMB = $SizeMB; Error = $Why }
}

$p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
if (-not $p) { return (New-DumpResult $null $null "pid $ProcessId is gone") }

if (-not $Path) {
    $dir = Join-Path $env:LOCALAPPDATA "CrashDumps"
    New-Item -ItemType Directory -Force $dir | Out-Null
    $Path = Join-Path $dir ("{0}.{1}.hang.dmp" -f $p.ProcessName, $ProcessId)
}

# full Namespace.Name, or the guard never matches and a second Add-Type throws
if (-not ("StretchXL.Dump" -as [type])) {
    Add-Type -Namespace StretchXL -Name Dump -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("dbghelp.dll", SetLastError = true)]
public static extern bool MiniDumpWriteDump(System.IntPtr hProcess, uint pid,
    Microsoft.Win32.SafeHandles.SafeFileHandle hFile, uint dumpType,
    System.IntPtr exceptionParam, System.IntPtr userStreamParam, System.IntPtr callbackParam);
"@ | Out-Null
}

# WithFullMemory(0x2) | WithHandleData(0x4) | WithUnloadedModules(0x20)
#   | WithProcessThreadData(0x100) | WithThreadInfo(0x1000)
$fullFlags = 0x2 -bor 0x4 -bor 0x20 -bor 0x100 -bor 0x1000
# WithIndirectlyReferencedMemory(0x40) keeps only memory the stacks point at.
# Not named $small: PowerShell names are case-insensitive, so that is the -Small switch.
$smallFlags = 0x4 -bor 0x20 -bor 0x40 -bor 0x100 -bor 0x1000
$type = if ($Small) { $smallFlags } else { $fullFlags }

$fs = [System.IO.File]::Create($Path)
try {
    $ok = [StretchXL.Dump]::MiniDumpWriteDump($p.Handle, [uint32]$ProcessId, $fs.SafeFileHandle,
                                               [uint32]$type, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero)
    $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
} finally { $fs.Close() }

if (-not $ok) {
    Remove-Item $Path -ErrorAction SilentlyContinue
    return (New-DumpResult $null $null ("MiniDumpWriteDump failed (win32 {0})" -f $err))
}
New-DumpResult $Path ([math]::Round((Get-Item $Path).Length / 1MB, 1)) $null
