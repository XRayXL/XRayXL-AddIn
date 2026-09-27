<#
  Turns the user docs into the web pages dist\docs\ ships: the Markdown GitHub shows, in
  xrayxl.com's look, readable from a folder with nothing but a browser.

    .\tools\Build-Docs.ps1 -Out build\docs     a preview: open build\docs\README.html

  tools\release.ps1 runs it into dist\docs\. Every local link and #anchor is checked once the
  pages are written, and a broken one stops it.
#>
# Not [Parameter(Mandatory)]: that empties $PSScriptRoot for the default below in Windows PowerShell.
param(
    [string]$Out,
    [string]$Root = (Split-Path $PSScriptRoot -Parent)
)
$ErrorActionPreference = 'Stop'
if (-not $Out) { throw 'Build-Docs.ps1 needs -Out <folder>' }
. (Join-Path $PSScriptRoot '_version.ps1')
. (Join-Path $PSScriptRoot '_markdig.ps1')

# The pages, in the order the navigation lists them.
$pages = @(
    @{ Source = 'README.md';               Page = 'README.html';           Nav = 'Start' }
    @{ Source = 'docs/DemoWalkthrough.md'; Page = 'DemoWalkthrough.html';  Nav = 'Walkthrough' }
    @{ Source = 'docs/TraceOptions.md';    Page = 'TraceOptions.html';     Nav = 'Trace options' }
    @{ Source = 'docs/TraceRowModel.md';   Page = 'TraceRowModel.html';    Nav = 'Trace rows' }
    @{ Source = 'docs/ErrorsInTheTrace.md';Page = 'ErrorsInTheTrace.html'; Nav = 'Errors' }
    @{ Source = 'docs/VBATracing.md';      Page = 'VBATracing.html';       Nav = 'VBA tracing' }
)
# Files at dist\'s root, one folder up from the pages.
$shippedAtRoot = @('LICENSE', 'THIRD-PARTY-NOTICES.txt')

$version = Get-XRayVersion -Root $Root
# Anything else a page links to is read on GitHub, as this release has it.
$github = "https://github.com/XRayXL/XRayXL-AddIn/blob/v$version/"

# Where a link in `source` points from dist\docs\: a page, an image, a file at dist\'s root, or GitHub.
function Resolve-Link([string]$source, [string]$href) {
    if ($href -match '^(#|[a-zA-Z][a-zA-Z0-9+.-]*:)') { return $href }        # an anchor, or absolute
    $url = [Uri]::new([Uri]::new($github + $source), $href)
    if (-not $url.AbsoluteUri.StartsWith($github)) { return $url.AbsoluteUri }  # above the tree, as ../../releases
    $path = [Uri]::UnescapeDataString($url.AbsolutePath.Substring(([Uri]$github).AbsolutePath.Length))
    $page = $pages | Where-Object { $_.Source -eq $path }
    if ($page) { return $page.Page + $url.Fragment }
    if ($path -like 'docs/images/*') { return 'images/' + $path.Substring('docs/images/'.Length) }
    if ($shippedAtRoot -contains $path) { return '../' + $path }
    if (-not (Test-Path -LiteralPath (Join-Path $Root $path))) { throw "$source links to $href, which is not in the repository" }
    return $github + $path + $url.Fragment
}

$template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light">
<title>{{TITLE}} &middot; XRayXL</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600&display=swap">
<style>
  /* xrayxl.com's palette and type, light only, as XRayXL-Perfetto.html */
  :root{
    --ground:#EFF1F3; --surface:#FFFFFF; --surface-2:#E5E9ED;
    --ink:#14181D; --ink-2:#4A535E; --ink-3:#78828E;
    --rule:#D3D8DE; --rule-strong:#B2BAC3;
    --signal:#C43C08; --steel:#1B5A78;
    --grid-line:rgba(20,24,29,.055);
    --display:"Archivo",system-ui,sans-serif;
    --body:"IBM Plex Sans",system-ui,-apple-system,"Segoe UI",sans-serif;
    --mono:"IBM Plex Mono",ui-monospace,"Cascadia Mono","Consolas",monospace;
  }
  *{box-sizing:border-box}
  body{
    margin:0; background:var(--ground); color:var(--ink);
    font-family:var(--body); font-size:16px; line-height:1.6;
    -webkit-font-smoothing:antialiased;
  }
  :focus-visible{outline:2px solid var(--signal); outline-offset:3px; border-radius:2px}
  .wrap{max-width:900px; margin-inline:auto; padding-inline:clamp(16px,5vw,48px)}

  /* ---------- heading and navigation, on the site's grid paper ---------- */
  .masthead{
    border-bottom:1px solid var(--rule);
    background:
      repeating-linear-gradient(to right, var(--grid-line) 0 1px, transparent 1px 88px),
      repeating-linear-gradient(to bottom, var(--grid-line) 0 1px, transparent 1px 30px);
  }
  .masthead .wrap{padding-block:22px 18px}
  .brandline{display:flex; align-items:center; gap:14px; margin-bottom:14px}
  .brand{font-family:var(--display); font-weight:800; font-size:24px; letter-spacing:-.035em; line-height:1}
  .brand span{color:var(--signal)}
  .tag{
    font-family:var(--mono); font-size:11px; letter-spacing:.14em; text-transform:uppercase;
    color:var(--signal); border:1px solid var(--signal); border-radius:2px; padding:3px 10px; line-height:1.3;
  }
  nav{display:flex; flex-wrap:wrap; gap:4px 22px; font-size:14.5px}
  nav a{color:var(--ink-2); text-decoration:none; padding-bottom:3px; border-bottom:2px solid transparent}
  nav a:hover{color:var(--signal)}
  nav a[aria-current]{color:var(--ink); font-weight:600; border-bottom-color:var(--signal)}

  /* ---------- the page ---------- */
  .doc{padding-block:34px 20px}
  .doc h1{font-family:var(--display); font-weight:800; font-size:clamp(30px,5vw,42px); letter-spacing:-.035em; line-height:1.1; margin:0 0 18px}
  .doc h2{font-family:var(--display); font-weight:700; font-size:26px; letter-spacing:-.02em; line-height:1.15;
          margin:44px 0 14px; padding-top:22px; border-top:1px solid var(--rule)}
  .doc hr + h2{margin-top:0; padding-top:0; border-top:none}
  .doc h3{font-family:var(--display); font-weight:700; font-size:19px; letter-spacing:-.01em; line-height:1.2; margin:30px 0 10px}
  .doc h4{font-size:16px; font-weight:600; margin:24px 0 8px}
  .doc p{margin:0 0 14px}
  .doc ul,.doc ol{margin:0 0 16px; padding-left:24px}
  .doc li{margin:4px 0}
  .doc li > p{margin-bottom:6px}
  .doc a{color:var(--steel); text-decoration-thickness:1px; text-underline-offset:2px}
  .doc a:hover{color:var(--signal)}
  .doc code{font-family:var(--mono); font-size:.88em; background:var(--surface-2); padding:1px 5px; border-radius:3px}
  .doc pre{background:var(--surface); border:1px solid var(--rule); border-radius:3px; padding:14px 16px;
           overflow-x:auto; margin:0 0 18px; line-height:1.5}
  .doc pre code{background:none; padding:0; font-size:12.5px}
  .doc .table{overflow-x:auto; margin:0 0 18px; background:var(--surface); border:1px solid var(--rule); border-radius:3px}
  .doc table{width:100%; border-collapse:collapse; font-size:14.5px}
  .doc th{background:var(--surface-2); color:var(--ink-3); font-family:var(--mono); font-weight:500; font-size:10.5px;
          letter-spacing:.1em; text-transform:uppercase; text-align:left; padding:9px 12px; border-bottom:1px solid var(--rule-strong)}
  .doc td{padding:8px 12px; border-bottom:1px solid var(--rule); vertical-align:top}
  .doc tr:last-child td{border-bottom:none}
  .doc blockquote{margin:0 0 18px; padding:14px 18px; background:var(--surface); border:1px solid var(--rule);
                  border-left:3px solid var(--signal); border-radius:3px; color:var(--ink-2)}
  .doc blockquote p:last-child{margin:0}
  .doc img{display:block; max-width:100%; height:auto; margin:6px 0 20px; border:1px solid var(--rule); border-radius:3px}
  .doc hr{border:0; border-top:1px solid var(--rule); margin:40px 0 30px}
  .doc details{margin:0 0 18px; padding:10px 16px; background:var(--surface); border:1px solid var(--rule); border-radius:3px}
  .doc summary{cursor:pointer; font-weight:600}

  footer{padding-block:18px 44px; border-top:1px solid var(--rule); font-size:13px; color:var(--ink-3)}
  footer a{color:var(--ink-3)}
</style>
</head>
<body>
<header class="masthead">
  <div class="wrap">
    <div class="brandline"><span class="brand">XRay<span>XL</span></span><span class="tag">Docs &middot; {{VERSION}}</span></div>
    <nav>
{{NAV}}
    </nav>
  </div>
</header>
<main class="wrap doc">
{{BODY}}
</main>
<footer class="wrap">XRayXL {{VERSION}} &middot; this page on GitHub: <a href="{{GITHUB}}">{{SOURCE}}</a></footer>
</body>
</html>
'@

$pipeline = Import-Markdig -Root $Root
New-Item -ItemType Directory -Force $Out | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)

foreach ($p in $pages) {
    $markdown = [IO.File]::ReadAllText((Join-Path $Root $p.Source))
    $body = [Markdig.Markdown]::ToHtml($markdown, $pipeline)
    # a wide table scrolls inside its frame, and a narrow one still fills it
    $body = [regex]::Replace($body, '<table(\s[^>]*)?>', '<div class="table">$0').Replace('</table>', '</table></div>')
    $body = [regex]::Replace($body, '\b(href|src)="([^"]*)"', {
        param($m)
        $target = Resolve-Link $p.Source ([Net.WebUtility]::HtmlDecode($m.Groups[2].Value))
        '{0}="{1}"' -f $m.Groups[1].Value, [Net.WebUtility]::HtmlEncode($target)
    })
    $h1 = [regex]::Match($body, '<h1[^>]*>(.*?)</h1>', 'Singleline')
    $title = if ($h1.Success) { [Net.WebUtility]::HtmlDecode(($h1.Groups[1].Value -replace '<[^>]+>', '')) } else { $p.Nav }
    $nav = ($pages | ForEach-Object {
        $current = if ($_.Page -eq $p.Page) { ' aria-current="page"' } else { '' }
        '      <a href="{0}"{1}>{2}</a>' -f $_.Page, $current, $_.Nav
    }) -join "`n"
    $html = $template.Replace('{{TITLE}}', [Net.WebUtility]::HtmlEncode($title)).Replace('{{VERSION}}', $version).
        Replace('{{NAV}}', $nav).Replace('{{GITHUB}}', $github + $p.Source).Replace('{{SOURCE}}', $p.Source).
        Replace('{{BODY}}', $body.TrimEnd())
    [IO.File]::WriteAllText((Join-Path $Out $p.Page), ($html -replace "`r`n", "`n"), $utf8)
}

$images = Join-Path $Out 'images'
New-Item -ItemType Directory -Force $images | Out-Null
Copy-Item (Join-Path $Root 'docs\images\*.png') $images -Force

# Every local link must reach a file, and every #anchor an id on its page. dist\'s root files are
# checked where they will be, beside $Out, only when they are already there.
$broken = @()
$ids = @{}
foreach ($p in $pages) {
    $text = [IO.File]::ReadAllText((Join-Path $Out $p.Page))
    $ids[$p.Page] = @([regex]::Matches($text, '\sid="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
}
foreach ($p in $pages) {
    $text = [IO.File]::ReadAllText((Join-Path $Out $p.Page))
    foreach ($m in [regex]::Matches($text, '\b(?:href|src)="([^"]*)"')) {
        $link = [Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
        if ($link -match '^[a-zA-Z][a-zA-Z0-9+.-]*:') { continue }
        $file, $anchor = $link -split '#', 2
        $target = if ($file) { $file } else { $p.Page }
        if ($file -and $file.StartsWith('../')) {
            if ((Test-Path -LiteralPath (Join-Path (Split-Path $Out) $file.Substring(3))) -or
                ($shippedAtRoot -contains $file.Substring(3))) { continue }
        }
        if (-not (Test-Path -LiteralPath (Join-Path $Out $target))) { $broken += "$($p.Page): $link -- no such file"; continue }
        if ($anchor -and $ids.ContainsKey($target) -and $ids[$target] -notcontains $anchor) {
            $broken += "$($p.Page): $link -- no heading or id '$anchor' on $target"
        }
    }
}
if ($broken) { throw ("broken links in the docs:`n  " + ($broken -join "`n  ")) }
Write-Host ("  docs -> {0}  ({1} pages)" -f $Out, $pages.Count)
