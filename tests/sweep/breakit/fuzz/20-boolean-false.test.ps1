. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxA(FALSE)";         W = 'boolean false'
           Value = '0'; Args = @('a1:A=0')
           Calls = @( @{ Function = 'TxA'; Args = 'a1:A=0'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
