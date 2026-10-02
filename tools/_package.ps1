# The release zip, XRayXL-<version>.zip: dist\'s files at its root, for anyone who would rather
# not clone. Dot-sourced by release.ps1.

# Zips the folder $Dist into $Out\XRayXL-<version>.zip, checks it, and returns its path.
function New-XRayPackage([string]$Dist, [string]$Version, [string]$Out) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    New-Item -ItemType Directory -Force $Out | Out-Null
    $zipPath = Join-Path (Resolve-Path -LiteralPath $Out).Path "XRayXL-$Version.zip"
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath }
    $root = (Resolve-Path -LiteralPath $Dist).Path.TrimEnd('\')
    $zip = [IO.Compression.ZipFile]::Open($zipPath, 'Create')
    try {
        Get-ChildItem -LiteralPath $root -Recurse -File | Sort-Object FullName | ForEach-Object {
            # forward slashes, which every unzipper reads as folders
            $name = $_.FullName.Substring($root.Length + 1).Replace('\', '/')
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $name, 'Optimal')
        }
    }
    finally { $zip.Dispose() }
    Assert-XRayPackage -ZipPath $zipPath -Version $Version
    return $zipPath
}

# Read back, every file the zip holds must be one its MANIFEST.txt lists, with the hash it lists,
# and every file the manifest lists must be there.
function Assert-XRayPackage([string]$ZipPath, [string]$Version) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $entries = @{}
        foreach ($e in $zip.Entries) { $entries[$e.FullName] = $e }
        if (-not $entries.ContainsKey('MANIFEST.txt')) { throw "$ZipPath has no MANIFEST.txt" }
        $reader = New-Object IO.StreamReader($entries['MANIFEST.txt'].Open())
        try { $manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }
        if (-not $manifest.StartsWith("XRayXL $Version`n")) { throw "the MANIFEST.txt in $ZipPath is not for XRayXL $Version" }

        $listed = @{}
        foreach ($m in [regex]::Matches($manifest, '(?m)^  ([0-9a-f]{64})  (.+)$')) {
            $listed[$m.Groups[2].Value.Replace('\', '/')] = $m.Groups[1].Value
        }
        foreach ($name in $entries.Keys) {
            if ($name -eq 'MANIFEST.txt') { continue }
            if (-not $listed.ContainsKey($name)) { throw "$ZipPath holds $name, which its MANIFEST.txt does not list" }
            $s = $entries[$name].Open()
            try { $hash = -join ($sha.ComputeHash($s) | ForEach-Object { $_.ToString('x2') }) } finally { $s.Dispose() }
            if ($hash -ne $listed[$name]) { throw "$name in $ZipPath is not the file its MANIFEST.txt hashed" }
        }
        $missing = @($listed.Keys | Where-Object { -not $entries.ContainsKey($_) })
        if ($missing) { throw "$ZipPath lacks $($missing -join ', '), which its MANIFEST.txt lists" }
    }
    finally { $sha.Dispose(); $zip.Dispose() }
}
