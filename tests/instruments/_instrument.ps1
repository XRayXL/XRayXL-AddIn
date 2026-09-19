# Shared helpers for the tests\ instruments (tracesoak, vbahammer), so their
# measurements stay comparable.

function Get-EnvInt {
    <#
      A positive whole number from the environment, or the default. Anything
      else throws rather than being coerced.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Default
    )
    $raw = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    $v = 0
    if (-not [int]::TryParse($raw, [ref]$v) -or $v -lt 1) {
        throw "$Name must be a positive whole number, got '$raw'"
    }
    return $v
}

function Enable-MultiThreadedCalc {
    <#
      Turn on multi-threaded calculation and return the thread count Excel
      chose (0 if it refused or the property is unavailable).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$App)
    try {
        $App.MultiThreadedCalculation.Enabled = $true
        $App.MultiThreadedCalculation.ThreadMode = 0      # automatic = one per core
        return [int]$App.MultiThreadedCalculation.ThreadCount
    } catch { return 0 }
}

function Get-ProcSample {
    <#
      One resource reading for the traced Excel, or $null if it is gone.
      Private bytes is the leak signal; the working set follows machine pressure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$ProcessId,
        # the caller's x-axis: cycles, arms, ...
        [int]$At = 0
    )
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    # GDI and USER objects are capped at 10,000 a process and are not in the handle count.
    if (-not ('XRayGui' -as [type])) {
        Add-Type -Name XRayGui -Namespace '' -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetGuiResources(System.IntPtr h, uint flags);'
    }
    $gdi = $null; $user = $null
    try { $gdi = [XRayGui]::GetGuiResources($p.Handle, 0); $user = [XRayGui]::GetGuiResources($p.Handle, 1) } catch {}
    [pscustomobject]@{
        At        = $At
        PrivateMB = [math]::Round($p.PrivateMemorySize64 / 1MB, 1)
        WorkingMB = [math]::Round($p.WorkingSet64 / 1MB, 1)
        Handles   = $p.HandleCount
        Gdi       = $gdi
        User      = $user
    }
}

function Set-XRayDepthAll {
    <#
      Both sources to DEPTH=ALL, with the echo checked. Returns '' when both
      took, otherwise what the setter actually said.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Sx)
    foreach ($source in @('VBA', 'XLL')) {
        $echo = Set-XRayTraceParam $Sx $source 'DEPTH' 'ALL'
        if ($echo -notmatch "$source DEPTH=ALL") { return "$source depth setter echo: $echo" }
    }
    return ''
}
