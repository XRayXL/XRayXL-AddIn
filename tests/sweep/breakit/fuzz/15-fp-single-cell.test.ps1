. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxK({0})";           W = 'FP, single cell'
           # TxK sums the cells.
           Calls = @( @{ Function = 'TxK'; Args = 'a1:K=Double[1..1,1..1]{0}'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
