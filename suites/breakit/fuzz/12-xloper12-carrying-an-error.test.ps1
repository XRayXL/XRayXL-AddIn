. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxQ(NA())";          W = 'XLOPER12 carrying an error' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
