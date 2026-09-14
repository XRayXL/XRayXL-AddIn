. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxOw({0})";          W = 'O-triple, single cell' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
