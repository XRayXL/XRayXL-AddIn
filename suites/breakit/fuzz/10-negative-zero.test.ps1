. (Join-Path $PSScriptRoot '..\_strings.ps1')
# the rendering of -0 is not pinned (0 and -0 are both honest); both arguments arriving is
$case = @{ F = "=TxB($negzero,0)";    W = 'negative zero'
           Value = '0'; ArgCount = 2; Args = @('a1:B=') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
