. (Join-Path $PSScriptRoot '..\_strings.ps1')
# TxDw returns the length; the rendered string is cut at a character and ends "...
$case = @{ F = "=TxDw($longWide)";    W = '400 wide chars, longer than the value buffer'
           Value = '400'; Args = @('a1:D%="WWWWWWWWWW', 'W"...') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
