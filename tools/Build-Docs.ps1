<#
  Turns the user docs into the web pages dist\docs\ ships: the Markdown GitHub shows, in
  xrayxl.com's look, readable from a folder with nothing but a browser. xrayxl.com serves
  the same pages in the same layout, so every link works from either.

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

# The pages, in the order the menu lists them and Previous and Next step through them.
$pages = @(
    @{ Source = 'README.md';               Page = 'README.html';           Nav = 'Start';         Blurb = 'What XRayXL is, and getting it running' }
    @{ Source = 'docs/DemoWalkthrough.md'; Page = 'DemoWalkthrough.html';  Nav = 'Walkthrough';   Blurb = 'Ten demo workbooks, a feature at a time' }
    @{ Source = 'docs/TraceOptions.md';    Page = 'TraceOptions.html';     Nav = 'Trace options'; Blurb = 'Every capture setting, and the log' }
    @{ Source = 'docs/TraceRowModel.md';   Page = 'TraceRowModel.html';    Nav = 'Trace rows';    Blurb = 'The trace file, column by column' }
    @{ Source = 'docs/ErrorsInTheTrace.md';Page = 'ErrorsInTheTrace.html'; Nav = 'Errors';        Blurb = 'How an error reads in a trace' }
    @{ Source = 'docs/VBATracing.md';      Page = 'VBATracing.html';       Nav = 'VBA tracing';   Blurb = 'How the VBA tracer works, in layers' }
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
  html{scroll-padding-top:76px}
  body{
    margin:0; background:var(--ground); color:var(--ink);
    font-family:var(--body); font-size:16px; line-height:1.6;
    -webkit-font-smoothing:antialiased;
  }
  :focus-visible{outline:2px solid var(--signal); outline-offset:3px; border-radius:2px}
  .wrap{max-width:900px; margin-inline:auto; padding-inline:clamp(16px,5vw,48px)}

  /* ---------- the bar along the top, on the site's grid paper, and its menu ---------- */
  .topbar{
    position:sticky; top:0; z-index:20; border-bottom:1px solid var(--rule);
    background-color:var(--ground);
    background-image:
      repeating-linear-gradient(to right, var(--grid-line) 0 1px, transparent 1px 88px),
      repeating-linear-gradient(to bottom, var(--grid-line) 0 1px, transparent 1px 30px);
  }
  .topbar .wrap{display:flex; align-items:center; gap:14px; min-height:60px}
  .brand{font-family:var(--display); font-weight:800; font-size:22px; letter-spacing:-.035em; line-height:1; color:var(--ink); text-decoration:none}
  .brand span{color:var(--signal)}
  .tag{
    font-family:var(--mono); font-size:11px; letter-spacing:.14em; text-transform:uppercase; white-space:nowrap;
    color:var(--signal); border:1px solid var(--signal); border-radius:2px; padding:3px 10px; line-height:1.3;
  }
  .menu{margin-left:auto; position:relative}
  .menu summary{
    list-style:none; cursor:pointer; display:flex; align-items:center; gap:10px; min-height:38px;
    padding:0 14px; background:var(--surface); border:1px solid var(--rule-strong); border-radius:3px;
    font-size:14px; font-weight:600; color:var(--ink);
  }
  .menu summary::-webkit-details-marker{display:none}
  .menu summary:hover,.menu[open] summary{border-color:var(--signal); color:var(--signal)}
  .bars,.bars::before,.bars::after{display:block; width:16px; height:2px; background:currentColor; transition:transform .15s, background-color .15s}
  .bars{position:relative}
  .bars::before,.bars::after{content:""; position:absolute; left:0}
  .bars::before{transform:translateY(-5px)}
  .bars::after{transform:translateY(5px)}
  .menu[open] .bars{background:transparent}
  .menu[open] .bars::before{transform:rotate(45deg)}
  .menu[open] .bars::after{transform:rotate(-45deg)}
  .menu-panel{
    position:absolute; right:0; top:calc(100% + 8px); width:min(330px, calc(100vw - 32px));
    max-height:calc(100vh - 84px); overflow-y:auto; padding:6px;
    background:var(--surface); border:1px solid var(--rule-strong); border-radius:4px;
    box-shadow:0 14px 34px rgba(20,24,29,.16);
  }
  .menu-head{
    font-family:var(--mono); font-size:10.5px; letter-spacing:.14em; text-transform:uppercase;
    color:var(--ink-3); padding:12px 12px 4px; border-top:1px solid var(--rule); margin-top:6px;
  }
  .menu-panel a{display:block; padding:7px 12px; border-radius:3px; text-decoration:none; color:var(--ink); font-size:14.5px; font-weight:600; line-height:1.35}
  .menu-panel a small{display:block; font-size:12.5px; font-weight:400; color:var(--ink-3)}
  .menu-panel a:hover{background:var(--surface-2); color:var(--signal)}
  .menu-panel a[aria-current]{background:var(--surface-2); box-shadow:inset 3px 0 0 var(--signal)}
  @media (prefers-reduced-motion:reduce){.bars,.bars::before,.bars::after{transition:none}}

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

  /* ---------- Previous and Next ---------- */
  .pager{display:grid; grid-template-columns:1fr 1fr; gap:12px; padding-block:30px 34px}
  .pager a{
    display:block; padding:13px 18px; background:var(--surface); border:1px solid var(--rule-strong); border-radius:3px;
    text-decoration:none; color:var(--ink); font-weight:600; line-height:1.35;
    transition:border-color .15s, color .15s;
  }
  .pager a:hover{border-color:var(--signal); color:var(--signal)}
  .pager small{display:block; font-family:var(--mono); font-size:10.5px; font-weight:500; letter-spacing:.12em; text-transform:uppercase; color:var(--ink-3); margin-bottom:3px}
  .pager .next{grid-column:2; text-align:right}
  @media (max-width:560px){.pager{grid-template-columns:1fr} .pager .next{grid-column:auto}}

  footer{padding-block:18px 44px; border-top:1px solid var(--rule); font-size:13px; color:var(--ink-3)}
  footer a{color:var(--ink-3)}
</style>
</head>
<body>
<header class="topbar">
  <div class="wrap">
    <a class="brand" href="https://xrayxl.com/">XRay<span>XL</span></a><span class="tag">Docs &middot; {{VERSION}}</span>
    <details class="menu">
      <summary><span class="bars" aria-hidden="true"></span>Menu</summary>
      <nav class="menu-panel" aria-label="Pages">
        <a href="https://xrayxl.com/">Home<small>xrayxl.com</small></a>
        <div class="menu-head">Guides</div>
{{NAV}}
        <div class="menu-head">Get it</div>
        <a href="https://github.com/XRayXL/XRayXL-AddIn/releases/download/v{{VERSION}}/XRayXL-{{VERSION}}.zip">Download<small>XRayXL-{{VERSION}}.zip</small></a>
        <a href="https://github.com/XRayXL/XRayXL-AddIn">View on GitHub<small>The source, and every release</small></a>
      </nav>
    </details>
  </div>
</header>
<main class="wrap doc">
{{BODY}}
</main>
<nav class="wrap pager" aria-label="Previous and next">
{{PAGER}}
</nav>
<footer class="wrap">XRayXL {{VERSION}} &middot; this page on GitHub: <a href="{{GITHUB}}">{{SOURCE}}</a></footer>
<script>
  // the menu also closes on Escape, or on a click outside it
  (function () {
    var menu = document.querySelector('.menu');
    document.addEventListener('click', function (e) { if (menu.open && !menu.contains(e.target)) menu.open = false; });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && menu.open) { menu.open = false; menu.querySelector('summary').focus(); }
    });
  })();
</script>
</body>
</html>
'@

$pipeline = Import-Markdig -Root $Root
New-Item -ItemType Directory -Force $Out | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)

for ($i = 0; $i -lt $pages.Count; $i++) {
    $p = $pages[$i]
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
        '        <a href="{0}"{1}>{2}<small>{3}</small></a>' -f $_.Page, $current, $_.Nav, $_.Blurb
    }) -join "`n"
    $pager = @()
    if ($i -gt 0) { $pager += '  <a class="prev" href="{0}"><small>&larr; Previous</small>{1}</a>' -f $pages[$i - 1].Page, $pages[$i - 1].Nav }
    if ($i -lt $pages.Count - 1) { $pager += '  <a class="next" href="{0}"><small>Next &rarr;</small>{1}</a>' -f $pages[$i + 1].Page, $pages[$i + 1].Nav }
    $html = $template.Replace('{{TITLE}}', [Net.WebUtility]::HtmlEncode($title)).Replace('{{VERSION}}', $version).
        Replace('{{NAV}}', $nav).Replace('{{PAGER}}', ($pager -join "`n")).Replace('{{GITHUB}}', $github + $p.Source).
        Replace('{{SOURCE}}', $p.Source).Replace('{{BODY}}', $body.TrimEnd())
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
