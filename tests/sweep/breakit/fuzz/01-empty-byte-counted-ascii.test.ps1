. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxD($empty)";        W = 'empty byte-counted ASCII'
           # TxD returns the count byte.
           Calls = @( @{ Function = 'TxD'; Args = 'a1:D=""'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
