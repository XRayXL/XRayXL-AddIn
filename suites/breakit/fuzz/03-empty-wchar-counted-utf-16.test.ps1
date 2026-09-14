. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxDw($empty)";       W = 'empty WCHAR-counted UTF-16' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
