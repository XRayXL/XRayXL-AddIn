. (Join-Path $PSScriptRoot '..\_strings.ps1')
# TxDw returns the length (16). The renderer escapes the quotes as \" and the CSV writer
# doubles every quote; the reader undoes only the doubling, so this is what it returns.
$case = @{ F = "=TxDw($quoted)";      W = 'commas and quotes -- CSV escaping'
           Value = '16'; Args = @('a1:D%="he said \"hi\", ok"') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
