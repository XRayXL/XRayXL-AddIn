. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxB($tiny,$huge)";   W = 'denormal-ish and near-max doubles'
           Value = '1E+307'; Args = @('a1:B=1e-15', 'a2:B=1e+307')
           # a*10 + b: 1e-14 vanishes against 1e+307.
           Calls = @( @{ Function = 'TxB'; Args = 'a1:B=1e-15 a2:B=1e+307'; Ret = '1e+307'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
