. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxOptional(5)";      W = 'omitted trailing argument' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
