. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxI(-32768)";        W = 'int16 minimum'
           Value = '-32768'; Args = @('a1:I=-32768')
           Calls = @( @{ Function = 'TxI'; Args = 'a1:I=-32768'; Ret = '-32768'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
