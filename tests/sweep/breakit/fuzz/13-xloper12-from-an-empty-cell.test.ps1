. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxQ(A1000)";         W = 'XLOPER12 from an EMPTY cell'
           # An empty cell arrives as xltypeNil, which TxQ does not name: 9. Its rendering is not specified.
           Calls = @( @{ Function = 'TxQ'; Ret = '9'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
