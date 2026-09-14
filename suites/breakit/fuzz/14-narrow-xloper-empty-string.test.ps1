. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxP($empty)";        W = 'narrow XLOPER, empty string' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
