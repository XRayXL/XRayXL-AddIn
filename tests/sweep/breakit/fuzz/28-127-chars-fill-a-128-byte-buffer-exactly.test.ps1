. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = '=TxC(REPT("a",127))'; W = '127 chars, the longest that fits the 128-byte read with its NUL'
           ArgsExact = 'a1:C="aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"'
           Calls = @( @{ Function = 'TxC'; Ret = '127'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
