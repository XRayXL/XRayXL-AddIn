. (Join-Path $PSScriptRoot '..\_strings.ps1')
# TxB is a*10+b, and 1e-14 vanishes against 1e307 at fifteen significant digits
$case = @{ F = "=TxB($tiny,$huge)";   W = 'denormal-ish and near-max doubles'
           Value = '1E+307'; Args = @('a1:B=1e-15', 'a2:B=1e+307') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
