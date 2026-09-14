. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxMany(1,2,3,4,5,6,7,8,9,10,11,12)"; W = 'TWELVE doubles: eight on the stack' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
