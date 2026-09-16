. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCw($empty)";       W = 'empty null-terminated UTF-16'
           # TxCw returns wcslen.
           Calls = @( @{ Function = 'TxCw'; Args = 'a1:C%=""'; Ret = '0'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
