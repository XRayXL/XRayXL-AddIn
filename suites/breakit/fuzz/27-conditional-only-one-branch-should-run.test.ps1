. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=IF(TxB(1,2)>0,TxB(5,6),TxB(7,8))"; W = 'conditional: only one branch should run' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
