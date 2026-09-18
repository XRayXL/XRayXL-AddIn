# The ribbon's transient registration, seen from outside. The CLSID is repeated here on purpose:
# read from the code under test, it could not catch the code changing it.
$script:RibbonProgId = 'XRayXL.RibbonUI'
$script:RibbonKeys = @(
    'HKCU:\Software\Classes\CLSID\{730E8300-14A0-45D3-AE1D-3EED7DDC3B30}'
    'HKCU:\Software\Classes\XRayXL.RibbonUI'
    'HKCU:\Software\Microsoft\Office\Excel\Addins\XRayXL.RibbonUI'
)

# The registry is per-user and the suite runs in parallel, so presence is sampled: only a key still
# there after every other session has had time to finish counts as left behind.
function Get-RibbonKeysLeftBehind([int]$Seconds = 12) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        $present = @($script:RibbonKeys | Where-Object { Test-Path $_ })
        if ($present.Count -eq 0) { return @() }
        Start-Sleep -Milliseconds 500
    }
    return @($script:RibbonKeys | Where-Object { Test-Path $_ })
}

# Is the add-in connected in this Excel? Asked of the object model, not inferred from the log.
function Get-RibbonConnected($App) {
    $coll = $null; $entry = $null
    try {
        $coll = $App.COMAddIns
        $entry = $coll.Item($script:RibbonProgId)
        return [bool]$entry.Connect
    }
    catch { return $false }
    finally {
        foreach ($o in @($entry, $coll)) {
            if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
        }
    }
}

# What the log says happened to the ribbon: 'loaded', 'off', or the fail-soft reason.
function Get-RibbonLogVerdict([string]$LogPath) {
    if (-not (Test-Path $LogPath)) { return 'no-log' }
    $t = Get-Content $LogPath -Raw
    if ($t -match 'ribbon: XRayXL buttons loaded') { return 'loaded' }
    if ($t -match 'XRAYXL_RIBBON=0')           { return 'off' }
    if ($t -match 'WARNING - ribbon: (.+?) --') { return 'failsoft: ' + $Matches[1] }
    return 'silent'
}

# Has Excel built the ribbon? onLoad fires only where there is a window, and these sessions are hidden.
function Test-RibbonUiLive([string]$LogPath) {
    if (-not (Test-Path $LogPath)) { return $false }
    return [bool]((Get-Content $LogPath -Raw) -match 'ribbon: controls are live')
}
