. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCallsBack(5)";     W = 'GENUINE nesting: add-in re-enters Excel via xlUDF' ; NestDepth = 2 }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
