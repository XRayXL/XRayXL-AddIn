. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxQ(A1000)";         W = 'XLOPER12 from an EMPTY cell' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
