. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCw($empty)";       W = 'empty null-terminated UTF-16' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
