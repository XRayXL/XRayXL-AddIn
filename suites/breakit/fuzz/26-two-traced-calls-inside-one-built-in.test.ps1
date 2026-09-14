. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=SUM(TxB(1,2),TxB(3,4))"; W = 'two traced calls inside one built-in' ; NestDepth = 1 }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
