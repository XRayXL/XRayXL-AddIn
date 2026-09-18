# One-shot install: copy the freshly built XLL into a single load-from folder. The add-in is one
# native XLL with no runtime dependencies beyond Windows.
#
# Point Excel's Add-ins list at build\addin\XRayXL64.xll (XRayXL32.xll for a 32-bit build) once;
# re-running this after a rebuild refreshes it in place. Close Excel first: a loaded XLL is
# locked.
param(
    [string]$Config  = "Release",
    [ValidateSet('x64', 'Win32')][string]$Platform = "x64",
    # defaulted below, so it follows -Root
    [string]$Target,
    [string]$Root    = (Split-Path $PSScriptRoot -Parent)
)
$ErrorActionPreference = "Stop"

if (-not $Target) { $Target = Join-Path $Root "build\addin" }

. (Join-Path $PSScriptRoot '_version.ps1')

$leaf = if ($Platform -eq 'Win32') { 'XRayXL32.xll' } else { 'XRayXL64.xll' }
$xll = Join-Path $Root "build\$Platform\$Config\XRayXL\$leaf"

if (-not (Test-Path $xll)) {
    throw "XLL not built: $xll  (build it first: msbuild src\XRayXL.vcxproj /p:Configuration=$Config /p:Platform=$Platform)"
}

# the installed binary must carry the version version.props declares
$version = Get-XRayVersion -Root $Root
Assert-XRayResourceVersion -Xll $xll -Version $version

# Refuse to run while the target XLL is loaded (Windows locks a mapped module).
if (Test-Path (Join-Path $Target $leaf)) {
    try {
        $s = [System.IO.File]::Open((Join-Path $Target $leaf), 'Open','ReadWrite','None'); $s.Close()
    } catch {
        # An Excel still has the XLL mapped. Killing it would leave a recovery
        # copy that the next Excel offers back behind a blocked-macros banner,
        # so ask our own hidden Excels to close their workbooks and quit first.
        & (Join-Path $PSScriptRoot "excel_shutdown.ps1")
        try {
            $s2 = [System.IO.File]::Open((Join-Path $Target $leaf), 'Open','ReadWrite','None'); $s2.Close()
        } catch {
            throw "$(Join-Path $Target $leaf) is still locked after asking Excel to close. An Excel that is not ours has it loaded -- close it, then re-run."
        }
    }
}

New-Item -ItemType Directory -Force $Target | Out-Null
Copy-Item $xll $Target -Force

$installed = Join-Path $Target $leaf
$ver = [Diagnostics.FileVersionInfo]::GetVersionInfo($installed).FileVersion
Write-Host ("  {0,-38} {1,8:N0} B   v{2}" -f (Split-Path $xll -Leaf), (Get-Item $xll).Length, $ver)

Write-Host "`ninstalled -> $Target"
Write-Host "In Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go > point it at:"
Write-Host "  $(Join-Path $Target $leaf)"
