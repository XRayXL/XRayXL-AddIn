. (Join-Path $PSScriptRoot '..\_strings.ps1')
# TxJ returns its argument
$case = @{ F = "=TxJ(-2147483648)";   W = 'int32 minimum'
           Value = '-2147483648'; Args = @('a1:J=-2147483648') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
