. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxB(TxB(1,2),3)";    W = 'cell-level nesting (sequential, not nested)' ; NestDepth = 1 }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
