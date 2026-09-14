. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCallsBack2(5)";    W = 'TWO levels of re-entry' ; NestDepth = 3 }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
