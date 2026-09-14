# Archives a release's PDBs with what identifies the builds they match. Dot-sourced by release.ps1.

function Get-XRayPdbSignature {
    <#
      The CodeView record a debugger matches a PDB by: the GUID and age, keyed as a
      symbol store keys them ("<GUID><age>"), and the PDB name the binary asks for.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $latin1 = [Text.Encoding]::GetEncoding(28591)   # one char per byte, so offsets match
    $bytes  = [IO.File]::ReadAllBytes($Path)
    $text   = $latin1.GetString($bytes)
    $at = 0
    while (($at = $text.IndexOf('RSDS', $at, [StringComparison]::Ordinal)) -ge 0) {
        $nameAt = $at + 24
        if ($nameAt -lt $bytes.Length) {
            $end  = $text.IndexOf([char]0, $nameAt)
            $name = if ($end -gt $nameAt) { $text.Substring($nameAt, $end - $nameAt) } else { '' }
            if ($name -match '\.pdb$') {
                $guidBytes = New-Object byte[] 16
                [Array]::Copy($bytes, $at + 4, $guidBytes, 0, 16)
                $age = [BitConverter]::ToUInt32($bytes, $at + 20)
                return [pscustomobject]@{
                    Key       = ('{0}{1:X}' -f ([Guid]::new($guidBytes)).ToString('N').ToUpperInvariant(), $age)
                    GuidBytes = $guidBytes
                    PdbName   = [IO.Path]::GetFileName($name)
                }
            }
        }
        $at += 4
    }
    return $null
}

function Save-XRaySymbols {
    <#
      Copies each binary's PDB, the release manifest and an index into <Archive>\<Tag>\.
      Everything is checked before anything is copied: a tag already archived is
      refused, and so is a PDB that does not carry its binary's GUID, because a PDB
      from another build symbolises nothing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Archive,
        [Parameter(Mandatory)][string]$Tag,
        [Parameter(Mandatory)][object[]]$Binaries,   # @{ Xll = <path>; Pdb = <path> }
        [Parameter(Mandatory)][string]$Manifest
    )

    $dest = Join-Path $Archive $Tag
    if ((Test-Path -LiteralPath $dest) -and (Get-ChildItem -LiteralPath $dest -Force | Select-Object -First 1)) {
        throw "symbols for $Tag are already archived at $dest"
    }
    if (-not (Test-Path -LiteralPath $Manifest)) { throw "no manifest: $Manifest" }

    $latin1  = [Text.Encoding]::GetEncoding(28591)
    $checked = foreach ($bin in $Binaries) {
        if (-not (Test-Path -LiteralPath $bin.Xll)) { throw "not built: $($bin.Xll)" }
        if (-not (Test-Path -LiteralPath $bin.Pdb)) { throw "no PDB: $($bin.Pdb)" }
        $sig = Get-XRayPdbSignature -Path $bin.Xll
        if (-not $sig) { throw "no CodeView record in $($bin.Xll) -- linked without debug information?" }
        if ($sig.PdbName -ine (Split-Path $bin.Pdb -Leaf)) {
            throw "$($bin.Xll) asks for $($sig.PdbName), not $(Split-Path $bin.Pdb -Leaf)"
        }
        $pdbText = $latin1.GetString([IO.File]::ReadAllBytes($bin.Pdb))
        if ($pdbText.IndexOf($latin1.GetString($sig.GuidBytes), [StringComparison]::Ordinal) -lt 0) {
            throw "$($bin.Pdb) is not the PDB for $($bin.Xll): it does not carry the binary's GUID"
        }
        [pscustomobject]@{ Xll = $bin.Xll; Pdb = $bin.Pdb; Sig = $sig }
    }

    New-Item -ItemType Directory -Force $dest | Out-Null
    $index = @("PDBs for $Tag. key is the GUID and age a debugger or symbol store matches by.", "")
    foreach ($c in $checked) {
        Copy-Item -LiteralPath $c.Pdb -Destination (Join-Path $dest $c.Sig.PdbName) -Force
        $index += ("{0,-20} sha256 {1}  pdb {2}  key {3}" -f (Split-Path $c.Xll -Leaf),
                   (Get-FileHash -LiteralPath $c.Xll -Algorithm SHA256).Hash.ToLower(),
                   $c.Sig.PdbName, $c.Sig.Key)
    }
    Copy-Item -LiteralPath $Manifest -Destination (Join-Path $dest 'MANIFEST.txt') -Force
    [IO.File]::WriteAllText((Join-Path $dest 'SYMBOLS.txt'), (($index -join "`n") + "`n"),
                            (New-Object Text.UTF8Encoding($false)))
    return $dest
}
