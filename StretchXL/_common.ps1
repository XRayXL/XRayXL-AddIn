# Shared by the manager, its job workers, the kit and tools\excel_shutdown.ps1, so every
# run finds, configures and kills Excel the same way. Job workers dot-source it by path.

# Guarded: the kit and the manager can both load this into one process.
if (-not ('StretchXL.WindowProcess' -as [type])) {
    Add-Type -Namespace StretchXL -Name WindowProcess -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr h, out int pid);
'@
}

function Get-WindowProcessId {
    <#
      The process id that owns a window, or 0 when the handle is not a window.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][IntPtr]$Hwnd)
    $owner = 0
    [void][StretchXL.WindowProcess]::GetWindowThreadProcessId($Hwnd, [ref]$owner)
    return $owner
}

function Invoke-WithRetry {
    <#
      Runs Action until it does not throw, up to Tries times, and returns its
      output. Returns nothing if every try threw.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [Parameter(Mandatory)][int]$Tries,
        [Parameter(Mandatory)][int]$DelayMs
    )
    for ($attempt = 0; $attempt -lt $Tries; $attempt++) {
        try { return (& $Action) } catch { Start-Sleep -Milliseconds $DelayMs }
    }
}

function Resolve-SuiteConfig {
    <#
      The suite configuration in force for one folder: every suite.psd1 from
      the tests root down to that folder, merged, nearest wins PER KEY, unset
      keys inherited from ancestors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SuiteDir,
        # the tests root, or the suite folder itself when standalone
        [Parameter(Mandatory)][string]$RootDir
    )
    $chain = New-Object System.Collections.ArrayList
    $walk = $SuiteDir
    while ($true) {
        [void]$chain.Insert(0, $walk)
        if ($walk.TrimEnd('\') -ieq $RootDir.TrimEnd('\')) { break }
        $parentDir = Split-Path $walk -Parent
        if (-not $parentDir -or $parentDir -eq $walk) { break }
        $walk = $parentDir
    }
    # keep climbing above the root while suite.psd1 files continue, so a
    # subfolder run inherits its suite's config
    $above = Split-Path $RootDir -Parent
    while ($above -and (Test-Path -LiteralPath (Join-Path $above 'suite.psd1'))) {
        [void]$chain.Insert(0, $above)
        $next = Split-Path $above -Parent
        if (-not $next -or $next -eq $above) { break }
        $above = $next
    }
    $merged = @{}
    foreach ($dir in $chain) {
        $cfgFile = Join-Path $dir 'suite.psd1'
        if (-not (Test-Path -LiteralPath $cfgFile)) { continue }
        $data = Import-PowerShellDataFile -LiteralPath $cfgFile
        foreach ($k in $data.Keys) {
            $v = $data[$k]
            if ($k -eq 'RegisterXll') {
                $v = @($v | ForEach-Object {
                    $p = [Environment]::ExpandEnvironmentVariables([string]$_)
                    if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $dir $p }
                    [System.IO.Path]::GetFullPath($p)
                })
            }
            elseif ($k -eq 'RequireNotOlderThan') {
                # both sides resolve from the declaring suite.psd1's folder
                $resolved = @{}
                foreach ($key in @($v.Keys)) {
                    $rk = [Environment]::ExpandEnvironmentVariables([string]$key)
                    if (-not [System.IO.Path]::IsPathRooted($rk)) { $rk = Join-Path $dir $rk }
                    # one reference, or a list when the file is built from several places
                    $resolved[[System.IO.Path]::GetFullPath($rk)] = @($v[$key] | ForEach-Object {
                        $rv = [Environment]::ExpandEnvironmentVariables([string]$_)
                        if (-not [System.IO.Path]::IsPathRooted($rv)) { $rv = Join-Path $dir $rv }
                        [System.IO.Path]::GetFullPath($rv)
                    })
                }
                $v = $resolved
            }
            $merged[$k] = $v
        }
    }
    return $merged
}

function Assert-SuiteArtifacts {
    <#
      Everything a resolved suite config claims about files, checked before a
      single Excel starts: the XLLs it registers exist, and nothing it
      registers is older than what it was built from.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SuiteDir,
        [Parameter(Mandatory)]$Config
    )
    foreach ($xllPath in @($Config['RegisterXll'])) {
        if ($xllPath -and -not (Test-Path -LiteralPath $xllPath)) {
            throw "suite config for '$SuiteDir' names an XLL that does not exist: $xllPath"
        }
    }
    # a stale artifact would pass green against an older build
    if (-not $Config.ContainsKey('RequireNotOlderThan')) { return }
    foreach ($have in @($Config['RequireNotOlderThan'].Keys)) {
        # the newest of every reference is the one the file must not be older than
        $ref = $null; $rT = $null
        foreach ($one in @($Config['RequireNotOlderThan'][$have])) {
            if (-not (Test-Path -LiteralPath $one)) {
                throw "suite config for '$SuiteDir': RequireNotOlderThan references a path that does not exist: $one"
            }
            if (-not (Test-Path -LiteralPath $have)) {
                throw "suite config for '$SuiteDir': '$have' does not exist, but '$one' does -- the install step has never run."
            }
            # a directory means its newest source file; build artefacts are ignored
            if (Test-Path -LiteralPath $one -PathType Container) {
                $newest = Get-ChildItem -LiteralPath $one -Recurse -File -ErrorAction SilentlyContinue |
                          Where-Object { $_.Extension -in '.cpp', '.h', '.c', '.hpp', '.asm', '.inc', '.def' } |
                          Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
                if (-not $newest) {
                    throw "suite config for '$SuiteDir': RequireNotOlderThan reference '$one' is a directory with no source files in it."
                }
                $oneT = $newest.LastWriteTimeUtc; $oneName = $newest.FullName
            }
            else { $oneT = (Get-Item -LiteralPath $one).LastWriteTimeUtc; $oneName = $one }
            if ($null -eq $rT -or $oneT -gt $rT) { $rT = $oneT; $ref = $oneName }
        }
        $hT = (Get-Item -LiteralPath $have).LastWriteTimeUtc
        if ($hT -lt $rT) {
            $msg  = "Stale artifact -- refusing to run (suite '$SuiteDir')."
            if ($ref -like '*\src\*') {
                $msg += "`n  A source file is newer than the binary: the build either failed or was never run."
            }
            $msg += "`n  registered    : $have"
            $msg += "`n                  written {0} UTC" -f $hT.ToString('yyyy-MM-dd HH:mm:ss')
            $msg += "`n  is older than : $ref"
            $msg += "`n                  written {0} UTC" -f $rT.ToString('yyyy-MM-dd HH:mm:ss')
            $msg += "`n  The tests would run against the older file and could report PASS on work it does not contain."
            $msg += "`n  Build, then install, then run."
            throw $msg
        }
    }
}

function Resolve-GroupBySession {
    <#
      Whether to keep same-session items together. On by default for a shuffled
      reuse run, where a full shuffle would throw the reuse away; -NoGroup turns
      that off, -Group asks for it in an unshuffled reuse run. Throws on a
      request that cannot be honoured.
    #>
    [CmdletBinding()]
    param(
        [switch]$Group,
        [switch]$NoGroup,
        [switch]$RandomOrder,
        [Parameter(Mandatory)][string]$SessionMode,
        [switch]$FloorMode
    )
    if ($Group -and $NoGroup) { throw '-GroupBySession and -NoGroupBySession contradict each other: give one' }
    if ($Group) {
        if ($FloorMode) { throw '-GroupBySession needs -Path: the floor has no suite set-ups to group' }
        if ($SessionMode -eq 'Fresh') { throw '-GroupBySession needs -SessionMode Reuse or ReuseClean: a fresh session is never shared' }
        return $true
    }
    if ($NoGroup) { return $false }
    return ($RandomOrder.IsPresent -and $SessionMode -ne 'Fresh' -and -not $FloorMode)
}

function Group-WorkBySession {
    <#
      The work list with items that share a session key kept together. Items
      keep their order within a group; the groups come in first-appearance
      order, or in an order shuffled by Rng. A session is only ever reused
      between items with one key, so interleaving keys buys nothing but restarts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IList]$Items,
        # the run's seeded generator, so a seed still replays the whole order
        [System.Random]$Rng
    )
    $keys = New-Object System.Collections.ArrayList
    $byKey = @{}
    foreach ($item in $Items) {
        $key = [string]$item.SessionKey
        if (-not $byKey.ContainsKey($key)) {
            $byKey[$key] = New-Object System.Collections.ArrayList
            [void]$keys.Add($key)
        }
        [void]$byKey[$key].Add($item)
    }
    if ($Rng) {
        for ($i = $keys.Count - 1; $i -gt 0; $i--) {   # Fisher-Yates
            $j = $Rng.Next($i + 1)
            $swap = $keys[$i]; $keys[$i] = $keys[$j]; $keys[$j] = $swap
        }
    }
    $grouped = New-Object System.Collections.ArrayList
    foreach ($key in $keys) { $grouped.AddRange($byKey[$key]) }
    return ,$grouped
}

function Stop-ExcelByPid {
    <#
      Kill one Excel, by pid and image name (pids are recycled). Returns
      $true if it killed something.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ProcessId)
    if ($ProcessId -le 0) { return $false }
    $victim = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if (-not $victim -or $victim.ProcessName -ne 'EXCEL') { return $false }
    try { $victim.Kill(); return $true } catch { return $false }
}
