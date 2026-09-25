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

# ---- Excel's events -------------------------------------------------------------------------

# A workbook the instruments poke to raise Excel's events: two sheets, VBA handlers of its own, and
# a formula only it depends on, so no model formula reads it.
$script:ScratchModule = @'
Public gN As Long
Public Function EV_Note(ByVal n As Long) As Long
    gN = gN + n
    EV_Note = gN
End Function
' From VBA, so the events arrive while traced frames are live; the write's Target has two areas.
Public Sub EV_Poke(ByVal i As Long)
    ThisWorkbook.Worksheets("E1").Range("B1:B3,D1:D3").Value = i
    ThisWorkbook.Worksheets("E1").Range("C" & (1 + i Mod 20)).Select
End Sub
'@
$script:ScratchSheet = @'
Private Sub Worksheet_Change(ByVal Target As Range)
    EV_Note Target.Count
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    EV_Note 1
End Sub
'@
function New-EventScratchBook {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Sx)
    New-XRayMacroBook $Sx 'EventScratch' -SheetName 'E1' -Components @(
        @{ Kind = 1; Name = 'EM'; Code = $script:ScratchModule }
        @{ Kind = 'Sheet'; Code = $script:ScratchSheet }
    ) -Cells @{ 'A1' = '=SUM(B1:B3,D1:D3)' } -Prepare {
        param($ws)
        $ws.Parent.Worksheets.Add([Type]::Missing, $ws).Name = 'E2'
    }
    return (Get-XRayMacroBook 'EventScratch')
}

# What one round raises: into the scratch book and back, a sheet change each way, a selection, an
# edit from outside VBA and one from inside, and the scratch sheet's recalculation.
$script:PokedEvents = @('WorkbookActivate', 'WorkbookDeactivate', 'WindowActivate', 'WindowDeactivate',
    'SheetActivate', 'SheetDeactivate', 'SheetSelectionChange', 'SheetChange', 'SheetCalculate', 'AfterCalculate')
function Invoke-EventPoke {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Sx, [Parameter(Mandatory)]$Scratch, [Parameter(Mandatory)]$ReturnTo,
          [Parameter(Mandatory)][int]$Round)
    $book = $Scratch.Book
    $book.Activate()
    $book.Worksheets.Item('E2').Activate()
    $book.Worksheets.Item('E1').Activate()
    [void]$Scratch.Sheet.Range(('A{0}' -f (2 + $Round % 20))).Select()
    $Scratch.Sheet.Range('B1').Value2 = [double]$Round
    [void]$Sx.App.Run(("'{0}'!EM.EV_Poke" -f $Scratch.Leaf), $Round)
    $ReturnTo.Activate()
    $book.Saved = $true       # a dirty book turns the close into a hidden save prompt
}

# Records exactly $On, changing only what differs from $Current (name -> bool), since each change is
# a call into Excel. Returns the echoes that did not agree.
function Set-EventChoice {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Sx, [string[]]$On = @(), [Parameter(Mandatory)][hashtable]$Current)
    $bad = @()
    foreach ($name in $script:AppEvents) {
        $want = $On -contains $name
        if ($Current.ContainsKey($name) -and $Current[$name] -eq $want) { continue }
        $echo = [string]$Sx.App.Run('XRayXL_SetTraceParam', 'EVENTS', $name, $want)
        if ($echo -notlike "EVENTS $name=$(if ($want) { 'TRUE' } else { 'FALSE' })*") { $bad += $echo }
        $Current[$name] = $want
    }
    return $bad
}

# Excel's event rows in a CSV trace, as name -> count. Streamed: a soak's file can run to gigabytes.
function Get-EventCounts {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Csv)
    $counts = @{}
    if (-not (Test-Path $Csv)) { return $counts }
    $rx = [regex]'^\d+,\d+,event,Excel,0,0,0,\d+,\d+,[^,]*,([^,]+),'
    foreach ($line in [System.IO.File]::ReadLines($Csv)) {
        $m = $rx.Match($line)
        if ($m.Success) { $counts[$m.Groups[1].Value] = 1 + [int]$counts[$m.Groups[1].Value] }
    }
    return $counts
}
