. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=IF(TxB(1,2)>0,TxB(5,6),TxB(7,8))"; W = 'conditional: only one branch should run'
           # The condition, then the branch it chose; TxB(7,8) must not appear at all.
           Calls = @(
               @{ Function = 'TxB'; Args = 'a1:B=1 a2:B=2'; Ret = '12'; Depth = '1'; Parent = -1 }
               @{ Function = 'TxB'; Args = 'a1:B=5 a2:B=6'; Ret = '56'; Depth = '1'; Parent = -1 }
           ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
