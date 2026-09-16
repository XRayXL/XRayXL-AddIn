. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxQ(NA())";          W = 'XLOPER12 carrying an error'
           # 4 is an error; an error value is written unquoted.
           Calls = @( @{ Function = 'TxQ'; Args = 'a1:Q=#N/A'; Ret = '4'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
