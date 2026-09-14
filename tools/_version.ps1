# The one version.props reader, dot-sourced by deploy.ps1 and release.ps1.

function Get-XRayVersion {
    <#
      The version version.props declares, as major.minor.patch. Read by name:
      there are several PropertyGroups, so positional access yields Object[].
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $propsPath = Join-Path $Root 'version.props'
    if (-not (Test-Path -LiteralPath $propsPath)) { throw "version.props not found at $propsPath" }

    $xml = [xml](Get-Content -LiteralPath $propsPath -Raw)
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', 'http://schemas.microsoft.com/developer/msbuild/2003')
    $part = {
        param($name)
        $n = $xml.SelectSingleNode("//m:$name", $ns)
        if (-not $n) { throw "version.props has no $name" }
        $n.InnerText.Trim()
    }
    $version = "{0}.{1}.{2}" -f (& $part 'XRayVersionMajor'),
                                (& $part 'XRayVersionMinor'),
                                (& $part 'XRayVersionPatch')
    if ($version -notmatch '^\d+\.\d+\.\d+$') {
        throw "version.props did not yield a version: '$version'"
    }
    return $version
}

function Assert-XRayResourceVersion {
    <#
      The XLL's VERSIONINFO must match version.props. The filename carries no
      version (Excel stores the path), so the resource is the only record.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Xll,
        [Parameter(Mandatory)][string]$Version
    )
    $resourceVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($Xll).FileVersion
    if ($resourceVersion -ne $Version) {
        throw ("version mismatch: version.props says $Version but $Xll carries $resourceVersion -- " +
               "rebuild, then re-run (a binary built before the bump would ship under a version it is not)")
    }
}
