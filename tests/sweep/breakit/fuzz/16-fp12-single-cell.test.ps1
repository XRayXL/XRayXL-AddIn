. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxKw({0})";          W = 'FP12, single cell'
           Calls = @( @{ Function = 'TxKw'; Args = 'a1:K%=Double[1..1,1..1]{0}'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
