. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxD($empty)";        W = 'empty byte-counted ASCII' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
