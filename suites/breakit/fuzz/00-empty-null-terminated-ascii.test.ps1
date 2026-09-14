. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxC($empty)";        W = 'empty null-terminated ASCII' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
