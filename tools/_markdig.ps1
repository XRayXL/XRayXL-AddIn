# Markdig, the Markdown library inside PowerShell 7's ConvertFrom-Markdown, for turning the docs into
# the HTML pages dist\ ships. Fetched from NuGet into build\tools on first use, pinned by version and
# SHA256, and loaded into this PowerShell. Nothing of it ships.

# Markdig's .NET Framework build and what it needs, in load order.
$script:MarkdigPackages = @(
    @{ Id = 'System.Runtime.CompilerServices.Unsafe'; Version = '6.1.2'; Sha256 = '5F6A7F53AF3465F92BEB6DA873EBE0E496206C313313B98BADEE4355A6B25937' }
    @{ Id = 'System.Buffers';                         Version = '4.6.1'; Sha256 = 'B00451E91D016FBEC091AD1E361F3A7015E1D91D4047F7E48A74455B2A673D79' }
    @{ Id = 'System.Numerics.Vectors';                Version = '4.6.1'; Sha256 = '2BC500A86DCB02F2032D6D877F9E2D6E9E4A79080E57239B4198679D4031F2C7' }
    @{ Id = 'System.Memory';                          Version = '4.6.3'; Sha256 = '26078AEB758C9AE985E8BF851F973026061DA6A5EB4837204D0C2D2204C72955' }
    @{ Id = 'Markdig';                                Version = '1.4.0'; Sha256 = 'BDD84353A3F499F989F3111B33ACEEBD9434F8A069CA8B3DA2A065F2230E7BAF' }
)

# Loads Markdig and returns a pipeline for GitHub's Markdown: pipe tables, GitHub's heading
# ids, so links to #anchors work as they do there, and bare URLs as links.
function Import-Markdig([string]$Root) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $cache = Join-Path $Root 'build\tools\markdig'
    New-Item -ItemType Directory -Force $cache | Out-Null
    # a loaded DLL cannot be replaced, and once loaded there is nothing to do
    $packages = if ('Markdig.Markdown' -as [type]) { @() } else { $script:MarkdigPackages }
    foreach ($p in $packages) {
        $dll = Join-Path $cache "$($p.Id).dll"
        $nupkg = Join-Path $cache ("{0}.{1}.nupkg" -f $p.Id, $p.Version)
        if (-not (Test-Path -LiteralPath $nupkg)) {
            $id = $p.Id.ToLower()
            Invoke-WebRequest "https://api.nuget.org/v3-flatcontainer/$id/$($p.Version)/$id.$($p.Version).nupkg" `
                -OutFile $nupkg -UseBasicParsing
        }
        $hash = (Get-FileHash -LiteralPath $nupkg -Algorithm SHA256).Hash
        if ($hash -ne $p.Sha256) {
            Remove-Item -LiteralPath $nupkg
            throw "$($p.Id) $($p.Version) from NuGet is not the package pinned here (SHA256 $hash)"
        }
        # the DLL is always taken afresh from the checked package
        $zip = [IO.Compression.ZipFile]::OpenRead($nupkg)
        try {
            $entry = $zip.Entries | Where-Object { $_.FullName -eq "lib/net462/$($p.Id).dll" }
            if (-not $entry) { throw "$($p.Id) $($p.Version) has no lib/net462 build" }
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dll, $true)
        }
        finally { $zip.Dispose() }
        Add-Type -Path $dll
    }
    $b = New-Object Markdig.MarkdownPipelineBuilder
    $b = [Markdig.MarkdownExtensions]::UsePipeTables($b)
    $b = [Markdig.MarkdownExtensions]::UseAutoIdentifiers($b, [Markdig.Extensions.AutoIdentifiers.AutoIdentifierOptions]::GitHub)
    $b = [Markdig.MarkdownExtensions]::UseAutoLinks($b)
    return $b.Build()
}
