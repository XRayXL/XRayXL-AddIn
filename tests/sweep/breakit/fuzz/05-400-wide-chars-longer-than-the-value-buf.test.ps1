. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxDw($longWide)";    W = '400 wide chars, longer than the value buffer'
           Value = '400'; Args = @('a1:D%="WWWWWWWWWW', 'W"...')
           # The count unit is 400; the rendered value is cut, so Args above pins its shape.
           Calls = @( @{ Function = 'TxDw'; Ret = '400'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
