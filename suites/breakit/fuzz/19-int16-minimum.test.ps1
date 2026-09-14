. (Join-Path $PSScriptRoot '..\_strings.ps1')
# without sign extension this reads 32768
$case = @{ F = "=TxI(-32768)";        W = 'int16 minimum'
           Value = '-32768'; Args = @('a1:I=-32768') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
