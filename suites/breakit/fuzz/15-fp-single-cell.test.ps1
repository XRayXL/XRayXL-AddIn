. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxK({0})";           W = 'FP, single cell' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
