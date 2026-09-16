. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxP($empty)";        W = 'narrow XLOPER, empty string'
           # TxP returns a type code: 2 is a string.
           Calls = @( @{ Function = 'TxP'; Args = 'a1:P=""'; Ret = '2'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
