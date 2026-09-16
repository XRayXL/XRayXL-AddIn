. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxOw({0})";          W = 'O-triple, single cell'
           Calls = @( @{ Function = 'TxOw'; Args = 'a1:O%=Double[1..1,1..1]{0}'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
