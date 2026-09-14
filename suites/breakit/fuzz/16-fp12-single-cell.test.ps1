. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxKw({0})";          W = 'FP12, single cell' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
