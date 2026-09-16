. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxDw($quoted)";      W = 'commas and quotes -- CSV escaping'
           Value = '16'; Args = @('a1:D%="he said \"hi\", ok"')
           Calls = @( @{ Function = 'TxDw'; Args = 'a1:D%="he said \"hi\", ok"'; Ret = '16'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
