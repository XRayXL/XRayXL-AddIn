. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxDw($empty)";       W = 'empty WCHAR-counted UTF-16'
           # TxDw returns the count unit.
           Calls = @( @{ Function = 'TxDw'; Args = 'a1:D%=""'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
