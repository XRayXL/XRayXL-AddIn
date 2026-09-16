. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxC($empty)";        W = 'empty null-terminated ASCII'
           # TxC returns strlen.
           Calls = @( @{ Function = 'TxC'; Args = 'a1:C=""'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
