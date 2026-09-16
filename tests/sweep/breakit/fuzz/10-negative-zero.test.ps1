. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxB($negzero,0)";    W = 'negative zero'
           Value = '0'; ArgCount = 2; Args = @('a1:B=')
           # -0*10 + 0 is +0. How -0 itself renders is not specified, so its argument is not pinned.
           Calls = @( @{ Function = 'TxB'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
