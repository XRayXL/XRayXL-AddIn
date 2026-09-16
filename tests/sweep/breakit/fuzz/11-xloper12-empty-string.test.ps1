. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxQ($empty)";        W = 'XLOPER12 empty string'
           # TxQ returns a type code: 2 is a string.
           Calls = @( @{ Function = 'TxQ'; Args = 'a1:Q=""'; Ret = '2'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
