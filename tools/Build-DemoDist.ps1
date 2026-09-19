# Assembles build\demo-dist\: a runnable copy of the current build, to try out. Not a release --
# dist\ is release.ps1's, after a green sweep. Needs no Excel; the demo workbooks come from dist\demo\.
#
#     .\tools\Build-DemoDist.ps1
param(
    [string]$Root = (Split-Path $PSScriptRoot -Parent),
    [string]$Configuration = 'Release'
)
$ErrorActionPreference = 'Stop'

$built = Join-Path $Root "build\x64\$Configuration"
$out   = Join-Path $Root 'build\demo-dist'
$outD  = Join-Path $out 'demo'
$srcD  = Join-Path $Root 'dist\demo'

$xll = Join-Path $built 'XRayXL\XRayXL64.xll'
if (-not (Test-Path $xll)) { throw "no build at $xll -- build the solution first" }

# a source file newer than the binary means the build is stale
$newest = Get-ChildItem (Join-Path $Root 'src') -Recurse -File |
          Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
if ($newest -and $newest.LastWriteTimeUtc -gt (Get-Item $xll).LastWriteTimeUtc) {
    throw ("$($newest.FullName) is newer than the built XLL -- rebuild first, " +
           "or the playground is not the code you just changed")
}

# Checked before anything is deleted: an Excel holding the XLL open would leave a wiped, half-built folder.
$existing = Join-Path $out 'XRayXL64.xll'
if (Test-Path $existing) {
    try { $h = [System.IO.File]::Open($existing, 'Open', 'Write', 'None'); $h.Close() }
    catch {
        $who = @(Get-Process EXCEL -EA SilentlyContinue |
                 ForEach-Object { "$($_.ProcessName) pid $($_.Id)" }) -join ', '
        throw ("$existing is locked, so nothing was changed. Close the Excel that has " +
               "the playground add-in loaded and run this again" +
               $(if ($who) { " -- running now: $who" } else { "" }))
    }
}

Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $outD | Out-Null

Copy-Item $xll $out -Force
foreach ($a in @('DemoFinance\DemoFinance64.xll', 'DemoBehaviors\DemoBehaviors64.xll')) {
    $p = Join-Path $built $a
    if (Test-Path $p) { Copy-Item $p $outD -Force } else { Write-Warning "missing: $p" }
}
Copy-Item (Join-Path $Root 'LICENSE')                 $out -Force
Copy-Item (Join-Path $Root 'docs\DemoWalkthrough.md') (Join-Path $outD 'README.md') -Force
# MinHook's licence wants its notice to travel with binaries; the committed file is current.
Copy-Item (Join-Path $Root 'THIRD-PARTY-NOTICES.txt') $out -Force

$books = @(Get-ChildItem $srcD -Filter *.xlsm -ErrorAction SilentlyContinue)
if (-not $books) { throw "no workbooks in $srcD -- run tools\Build-DemoWorkbooks.ps1 (it needs Excel)" }
$books | Copy-Item -Destination $outD -Force

# Say what this is and which commit it came from, so a stray copy can be identified.
$sha    = (& git -C $Root rev-parse --short HEAD 2>$null)
$branch = (& git -C $Root rev-parse --abbrev-ref HEAD 2>$null)
$dirty  = if ((& git -C $Root status --porcelain 2>$null)) { ' plus uncommitted changes' } else { '' }
$playground = @"
# XRayXL playground

**This is not a release.** ``dist\`` is untouched -- a release is assembled only by
``tools\release.ps1``, and only after a full green sweep. This was assembled by
``tools\Build-DemoDist.ps1`` from ``build\x64\$Configuration\`` so the build can be
tried out. There is no ``MANIFEST.txt``, because nothing here has been certified.

Built from ``$branch`` at ``$sha``$dirty, on $(Get-Date -Format 'yyyy-MM-dd HH:mm').

Open Excel, then drag ``XRayXL64.xll`` onto it (or double-click it). The **XRayXL**
group appears at the far right of the Developer tab once a workbook is open;
``demo\`` holds the workbooks and the two demo add-ins they use.
"@
# no BOM: PowerShell 5.1's -Encoding utf8 writes one
[System.IO.File]::WriteAllText((Join-Path $out 'PLAYGROUND.md'), $playground,
                               (New-Object System.Text.UTF8Encoding($false)))

"demo-dist assembled at $out"
Get-ChildItem $out -Recurse -File | ForEach-Object {
    "  {0,-34} {1,8:N0}" -f $_.FullName.Substring($out.Length + 1), $_.Length
}
