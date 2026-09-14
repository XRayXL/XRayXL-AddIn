# Assemble dist\ -- the committed, downloadable build -- and optionally publish it.
#
# WHY THE BINARIES ARE COMMITTED. dist\ is what someone gets by cloning, or by
# taking GitHub's automatic "Source code" archive from a tag: a working tool and
# a runnable demo, with no compiler. That is the whole point of it, and it is
# why the release needs no uploaded assets.
#
# THE RULE THAT KEEPS IT HONEST: only this script writes dist\. A normal build
# never does, and neither does the test process -- the suites load from
# build\addin\ and build\x64\Release\, never from here. So dist\ does not drift with
# the working tree: it lags it, and MANIFEST.txt says by exactly how much.
#
#     .\tools\release.ps1                          build, sweep, assemble dist\  (stops there)
#     .\tools\release.ps1 -Publish -Remote <name>  ...then commit, tag, push to <name>, release there
#     ... -SymbolArchive <dir>                     ...and first keep the PDBs under <dir>\<tag>\
#
# The sweep is the gate, and it runs HERE rather than in CI because GitHub's
# runners have MSBuild but no Excel: a CI-built release would be an untested
# binary, which is the one thing this project will not ship.
param(
    [switch]$Publish,
    # The git remote a publish pushes and releases to. Required with -Publish; no default.
    [string]$Remote,
    # Where to keep each release's PDBs, under a folder named by tag. Publish only.
    [string]$SymbolArchive,
    [string]$Root = (Split-Path $PSScriptRoot -Parent),
    [int]$Parallel = 8
)
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot '_version.ps1')
. (Join-Path $PSScriptRoot '_symbols.ps1')

function Write-Step($n, $what) { Write-Host "`n[$n] $what" -ForegroundColor Cyan }

# Where a publish goes: the remote's URL, and the GitHub owner/repo that gh must target too.
function Resolve-XRayRemote([string]$Name) {
    $url = & git -C $Root config --get "remote.$Name.url"
    if ($LASTEXITCODE -ne 0 -or -not $url) { throw "no git remote named '$Name' in $Root" }
    $url = "$url".Trim()
    if ($url -notmatch 'github\.com[:/](?<repo>[^/]+/[^/]+?)(\.git)?/?$') {
        throw "remote '$Name' ($url) is not a GitHub repository -- gh cannot create a release there"
    }
    [pscustomobject]@{ Url = $url; Repo = $Matches['repo'] }
}

$dist    = Join-Path $Root 'dist'
$distDemo= Join-Path $dist 'demo'
$built   = Join-Path $Root 'build\x64\Release'

# WHAT COUNTS AS A MODIFIED TREE. dist\ is excluded on purpose: this script
# writes it, so it is always modified by the time anyone could look.
# Uncompiled files count too: LICENSE and README ship, and the .sln decides what builds.
$srcPaths = @('src', 'StretchXL', 'suites', 'tests', 'tools',
              'docs', 'version.props', 'LICENSE', 'README.md', 'XRayXL.sln')
function Test-TreeDirty { [bool](& git -C $Root status --porcelain -- @srcPaths) }

# ---------------------------------------------------------------------------
# 1. Version. version.props is the single source; the binary must agree with it.
# ---------------------------------------------------------------------------
Write-Step 1 "Version"
$version = Get-XRayVersion -Root $Root
$tag = "v$version"
Write-Host "  version.props -> $version  (tag $tag)"

if ($SymbolArchive -and -not $Publish) {
    Write-Host "  -SymbolArchive: a dry run ships nothing, so no PDBs are archived"
}
if ($Remote -and -not $Publish) {
    $dest = Resolve-XRayRemote $Remote
    Write-Host ("  -Remote: a publish would go to {0}  ({1})" -f $dest.Url, $dest.Repo)
}

if ($Publish) {
    # NAMED ON EVERY RUN, NEVER ASSUMED: a publish pushes and creates a release,
    # and whatever `origin` happens to be in this clone is not a safe default.
    if (-not $Remote) { throw "-Publish needs -Remote <name>: the git remote to push and release to" }
    $dest = Resolve-XRayRemote $Remote
    Write-Host ("  publishing to: {0}  ({1})" -f $dest.Url, $dest.Repo)
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh CLI not found" }
    # A tag that already exists would publish a release pointing at OLD code.
    if ((& git -C $Root tag --list $tag)) {
        throw "tag $tag already exists -- bump version.props before publishing again"
    }
    $remoteTag = & git -C $Root ls-remote --tags $Remote "refs/tags/$tag"
    if ($LASTEXITCODE -ne 0) { throw "could not reach remote '$Remote' ($($dest.Url))" }
    if ($remoteTag) { throw "tag $tag already exists on $($dest.Repo) -- bump version.props before publishing again" }
    if ($SymbolArchive -and (Test-Path -LiteralPath (Join-Path $SymbolArchive $tag))) {
        throw "symbols for $tag are already archived under $SymbolArchive"
    }
    # REFUSED FROM A MODIFIED TREE. A dry run may be dirty -- that is what a dry
    # run is for, and the manifest says so. A PUBLISH may not: it tags, and a
    # tag whose manifest reads "assembled from a MODIFIED working tree" names a
    # release that nobody, the author included, can rebuild from that tag.
    # Checked HERE, before the ten minutes of build and sweep, not after.
    if (Test-TreeDirty) {
        throw ("refusing to publish from a modified working tree -- commit first, " +
               "then re-run. Uncommitted:" + [Environment]::NewLine +
               ((& git -C $Root status --porcelain -- @srcPaths) -join [Environment]::NewLine))
    }
}

# ---------------------------------------------------------------------------
# 2. Build everything, from clean output, so nothing stale can be copied.
# ---------------------------------------------------------------------------
Write-Step 2 "Build"
# REQUIRE THE C++ TOOLSET, not merely MSBuild. A Community install without the
# Desktop-development-with-C++ workload ships MSBuild and no compiler, and it
# can easily be the "latest" install -- asking for Microsoft.Component.MSBuild
# selects it and then fails on the v145 toolset. Ask for what we actually need.
$msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
            -latest -products * `
            -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
            -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuild) { throw "No Visual Studio install with the C++ toolset (Desktop development with C++)" }
Write-Host "  msbuild: $msbuild"

& $msbuild (Join-Path $Root 'XRayXL.sln') /p:Configuration=Release /p:Platform=x64 /t:Rebuild /v:m /nologo
if ($LASTEXITCODE -ne 0) { throw "build failed" }

# The resource is the only thing that says which build a binary is. If it
# disagrees with version.props, this XLL predates the last bump.
$productXll = Join-Path $built 'XRayXL\XRayXL64.xll'
Assert-XRayResourceVersion -Xll $productXll -Version $version

# ---------------------------------------------------------------------------
# 3. Deploy, then sweep. The suites are the gate.
# ---------------------------------------------------------------------------
Write-Step 3 "Deploy"
& (Join-Path $PSScriptRoot 'deploy.ps1')   # throws on failure; a script call sets no exit code

Write-Step 4 "Full sweep (this takes several minutes and drives real Excel)"
$sweepOut = Join-Path $env:TEMP ("XRayXL-release-$version-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
& (Join-Path $Root 'StretchXL\StretchXL.ps1') -Parallel $Parallel `
    -Path (Join-Path $Root 'suites') -OutDir $sweepOut
$sweepExit = $LASTEXITCODE
if ($sweepExit -ne 0) {
    throw "SWEEP FAILED (exit $sweepExit) -- nothing assembled. Results: $sweepOut"
}
Write-Host "  sweep green -> $sweepOut"

# ---------------------------------------------------------------------------
# 4b. PROVE THE SWEEP TESTED WHAT WE ARE ABOUT TO SHIP.
#
# "The sweep exited 0" and "this file is good" are different claims, and only
# this script's ordering connects them. That is not evidence: a rebuild, a
# stray deploy or a second shell between the two steps breaks the link
# silently, and the manifest would still say PASS. So read the identity the
# sweep recorded and compare hashes.
# ---------------------------------------------------------------------------
$resultsFile = Get-ChildItem (Join-Path $sweepOut 'results-*.jsonl') |
               Sort-Object LastWriteTime | Select-Object -Last 1
if (-not $resultsFile) { throw "sweep produced no results file in $sweepOut" }

$meta = Get-Content -LiteralPath $resultsFile.FullName -TotalCount 1 | ConvertFrom-Json
if ($meta.type -ne 'meta') { throw "first line of $($resultsFile.Name) is not the metadata record" }
if (-not $meta.product)    { throw "the sweep recorded no product block -- StretchXL is older than this script expects" }

$debugTested = @($meta.product | Where-Object { $_.isDebug })
if ($debugTested) {
    throw ("the sweep tested a DEBUG build: {0}. A release must be Release throughout." -f
           (($debugTested | ForEach-Object { $_.name }) -join ', '))
}

$shipHash  = (Get-FileHash -LiteralPath $productXll -Algorithm SHA256).Hash.ToLower()
$testedXll = @($meta.product | Where-Object { $_.name -ieq 'XRayXL64.xll' })
if ($testedXll.Count -eq 0) { throw "the sweep's product block does not mention XRayXL64.xll" }
if ($testedXll.Count -gt 1) {
    throw ("the sweep registered {0} different XRayXL64.xll files -- cannot say which one passed" -f $testedXll.Count)
}
if ($testedXll[0].sha256 -ne $shipHash) {
    throw ("THE SWEEP DID NOT TEST THIS BINARY. Tested {0} ({1}); about to ship {2} ({3}). " -f
           $testedXll[0].sha256.Substring(0,12), $testedXll[0].path,
           $shipHash.Substring(0,12), $productXll) +
          "Something rebuilt or redeployed during the run; re-run release.ps1."
}
Write-Host ("  sweep tested the binary being shipped: sha {0} v{1} release" -f `
            $shipHash.Substring(0,12), $testedXll[0].fileVersion)

# ---------------------------------------------------------------------------
# 5. Assemble dist\. Replace it wholesale: a file left behind from a previous
#    version is exactly the stale copy this layout exists to avoid.
# ---------------------------------------------------------------------------
# A dry run does not touch dist\: a check that modified a committed folder would
# leave the tree dirty. It assembles into a scratch folder beside the sweep
# results instead, so what would ship can be inspected, manifest and all.
$staging   = -not $Publish
$stage     = if ($staging) { Join-Path $sweepOut 'dist-preview' } else { $dist }
$stageDemo = Join-Path $stage 'demo'

Write-Step 5 $(if ($staging) { "Assemble a preview (dry run -- dist\ is not touched)" } else { "Assemble dist\" })

# The workbooks are generated by a separate, Excel-dependent step and are never
# rebuilt here, so they are the one thing taken FROM dist\ rather than written
# to it -- in a dry run they are copied into the preview so its manifest hashes
# the same set of files a release would.
$workbooks = @(Get-ChildItem (Join-Path $distDemo '*.xlsm') -ErrorAction SilentlyContinue)
if (-not $workbooks) {
    throw "no workbooks in dist\demo\ -- run tools\Build-DemoWorkbooks.ps1 first (it needs Excel)"
}
if ($staging) {
    New-Item -ItemType Directory -Force $stageDemo | Out-Null
    $workbooks | Copy-Item -Destination $stageDemo -Force
}

# Everything that is not a workbook is replaced.
$doomed = @(Get-ChildItem $stage -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -ne '.xlsm' })

# Check every file before deleting any of it. An Excel with dist\XRayXL64.xll
# loaded makes that file undeletable, and deleting as we go could leave dist\
# without its licence and manifest but still holding the old binary.
$locked = @()
foreach ($d in $doomed) {
    try { $s = [System.IO.File]::Open($d.FullName, 'Open', 'ReadWrite', 'None'); $s.Close() }
    catch { $locked += $d.FullName }
}
if ($locked) {
    throw ("dist\ is in use and NOTHING has been changed. Close whatever has these loaded " +
           "-- an Excel with the add-in or a demo workbook open will do it -- then re-run:" +
           [Environment]::NewLine + "  " + ($locked -join ([Environment]::NewLine + "  ")))
}

$doomed | Remove-Item -Force
New-Item -ItemType Directory -Force $stageDemo | Out-Null

Copy-Item $productXll $stage -Force
Copy-Item (Join-Path $built 'DemoFinance\DemoFinance64.xll')     $stageDemo -Force
Copy-Item (Join-Path $built 'DemoBehaviors\DemoBehaviors64.xll') $stageDemo -Force
Copy-Item (Join-Path $Root 'LICENSE')          $stage -Force
Copy-Item (Join-Path $Root 'docs\DemoWalkthrough.md') (Join-Path $stageDemo 'README.md') -Force

# LF before hashing: git stores these LF, so a CRLF copy would fail its own manifest
foreach ($textFile in @((Join-Path $stage 'LICENSE'), (Join-Path $stageDemo 'README.md'))) {
    $raw = [System.IO.File]::ReadAllText($textFile)
    if ($raw.Contains("`r`n")) {
        [System.IO.File]::WriteAllText($textFile, ($raw -replace "`r`n", "`n"),
                                       (New-Object System.Text.UTF8Encoding($false)))
        Write-Host ("  normalised to LF: {0}" -f (Split-Path $textFile -Leaf))
    }
}

# MinHook's BSD-2-Clause requires its notice to travel with BINARY
# redistributions, not only source. dist\ is a binary redistribution.
$notices = @()
$notices += "THIRD-PARTY NOTICES for XRayXL $version"
$notices += ("=" * 60)
$notices += ""
$notices += "XRayXL itself is licensed under the GNU GPL v3 -- see LICENSE."
$notices += "It includes the following components under their own terms."
$notices += ""
$notices += ("-" * 60)
$notices += "MinHook -- the inline hooking library"
$notices += ("-" * 60)
$notices += ""
$notices += (Get-Content (Join-Path $Root 'src\third_party\minhook\LICENSE.txt') -Raw).TrimEnd()
$notices += ""
$notices += ("-" * 60)
$notices += "xlcall.h -- the Excel C API header"
$notices += ("-" * 60)
$notices += ""
$notices += "From Microsoft's Excel XLL SDK, used under Microsoft's SDK terms."
$notices += "The header and its provenance note are in src\third_party\xlcall.h."
# LF, so the hashed bytes match the committed ones
[System.IO.File]::WriteAllText(
    (Join-Path $stage 'THIRD-PARTY-NOTICES.txt'),
    (($notices -join "`n") + "`n"),
    (New-Object System.Text.UTF8Encoding($false)))

# ---------------------------------------------------------------------------
# 6. MANIFEST.txt -- what this dist IS. Without it a committed binary is a file
#    of unknown provenance, which is the fair objection to committing one.
# ---------------------------------------------------------------------------
Write-Step 6 "Manifest"
# the source commit: the commit that adds this file cannot know its own sha
$commit = (& git -C $Root rev-parse HEAD).Trim()
$dirty  = Test-TreeDirty
$lines = @()
$lines += "XRayXL $version"
$lines += ""
$lines += "tag        : $tag$(if ($staging) { '  (a dry run creates no tag)' } else { '  (on the commit that adds this folder)' })"
$lines += "commit     : $commit"
$lines += "             the SOURCE commit this was built and swept from. The commit that"
$lines += "             ADDS this folder is its child, and the tag is on that child."
$lines += "built      : {0:yyyy-MM-dd HH:mm:ss zzz}" -f (Get-Date)
$lines += "sweep      : PASS ($(Split-Path $sweepOut -Leaf))"
$lines += "verified   : the sweep's recorded SHA256 for XRayXL64.xll matches the file shipped here"
if ($dirty) { $lines += "WARNING    : assembled from a MODIFIED working tree, not a clean $commit" }
$lines += ""
$lines += "This folder is assembled only by tools\release.ps1. A normal build does"
$lines += "not write it, so between releases it lags the source it sits beside."
$lines += ""
$lines += "SHA256:"
Get-ChildItem $stage -Recurse -File |
    Where-Object { $_.Name -ne 'MANIFEST.txt' } |
    Sort-Object FullName | ForEach-Object {
        $rel = $_.FullName.Substring($stage.Length + 1)
        $lines += ("  {0}  {1}" -f (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower(), $rel)
    }
# LF, as for the notices
[System.IO.File]::WriteAllText(
    (Join-Path $stage 'MANIFEST.txt'),
    (($lines -join "`n") + "`n"),
    (New-Object System.Text.UTF8Encoding($false)))

Write-Host ""
Get-ChildItem $stage -Recurse -File | Sort-Object FullName | ForEach-Object {
    Write-Host ("  {0,-34} {1,9:N0} B" -f $_.FullName.Substring($stage.Length + 1), $_.Length)
}

if ($staging) {
    Write-Host "`nDRY RUN for $version -- dist\ was NOT touched." -ForegroundColor Green
    Write-Host "What a release would ship is staged for inspection at:"
    Write-Host "  $stage"
    Write-Host "Commit your changes, then re-run with -Publish to write dist\ for real."
    return
}

# ---------------------------------------------------------------------------
# 7. Keep the PDBs, before anything is committed: a release whose symbols could
#    not be kept does not go out. They stay out of dist\ -- each is several MB
#    and records the build machine's paths.
# ---------------------------------------------------------------------------
if ($SymbolArchive) {
    Write-Step 7 "Archive symbols"
    $obj = Join-Path $Root 'build\obj\x64\Release'
    $saved = Save-XRaySymbols -Archive $SymbolArchive -Tag $tag -Manifest (Join-Path $stage 'MANIFEST.txt') -Binaries @(
        @{ Xll = $productXll;                                          Pdb = (Join-Path $obj 'XRayXL\XRayXL64.pdb') }
        @{ Xll = (Join-Path $built 'DemoFinance\DemoFinance64.xll');     Pdb = (Join-Path $obj 'DemoFinance\DemoFinance64.pdb') }
        @{ Xll = (Join-Path $built 'DemoBehaviors\DemoBehaviors64.xll'); Pdb = (Join-Path $obj 'DemoBehaviors\DemoBehaviors64.pdb') }
    )
    Write-Host "  PDBs -> $saved  (commit them where that folder lives)"
}

# ---------------------------------------------------------------------------
# 8. Publish. Commit dist\, tag, and let GitHub generate the source archives --
#    which now CONTAIN the binaries, so there are no assets to upload.
# ---------------------------------------------------------------------------
Write-Step 8 "Publish"
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh CLI not found" }

# The dist\ commit is the release ledger: "git log --oneline -- dist/" reads as
# one line per release, saying which binary shipped and what gated it. (It is a
# second commit because the manifest hashes binaries built from the first.)
$resultCount = 0
foreach ($line in [System.IO.File]::ReadLines($resultsFile.FullName)) { $resultCount++ }
$resultCount = [Math]::Max(0, $resultCount - 1)      # the metadata line is not a result

$msg = @()
$msg += "Release $tag -- XRayXL64.xll $($shipHash.Substring(0,12)), $resultCount tests green"
$msg += ""
$msg += "binary     : XRayXL64.xll sha256 $shipHash"
$msg += "built from : $commit (this commit's parent -- the source that was swept)"
$msg += "sweep      : $resultCount result(s), all benign -- $(Split-Path $sweepOut -Leaf)"
$msg += "verified   : the sweep's recorded SHA256 matches the binary committed here"
$msg += ""
$msg += "dist\ is assembled only by tools\release.ps1 and only after that sweep"
$msg += "passes. MANIFEST.txt carries the SHA256 of every file in it."

& git -C $Root add -- dist
& git -C $Root commit -m ($msg -join [Environment]::NewLine)
if ($LASTEXITCODE -ne 0) { throw "commit failed" }
# without a tag, "gh release create" would invent one at the remote branch tip
& git -C $Root tag -a $tag -m "XRayXL $version"
if ($LASTEXITCODE -ne 0) { throw "tag $tag failed -- dist\ is committed but nothing is tagged or pushed" }
& git -C $Root push $Remote HEAD --follow-tags
if ($LASTEXITCODE -ne 0) { throw "push failed" }

# No assets: dist\ is in the tree, so GitHub's generated source archives already
# contain the built tool and a runnable demo.
& gh release create $tag --repo $($dest.Repo) --title "XRayXL $version" --generate-notes
if ($LASTEXITCODE -ne 0) { throw "gh release create failed" }

Write-Host "`npublished $tag" -ForegroundColor Green
