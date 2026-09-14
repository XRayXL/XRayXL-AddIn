. (Join-Path $PSScriptRoot '..\_strings.ps1')
# the newline renders as \n inside the quotes, so the record stays on one line
$case = @{ F = "=TxDw($newline)";     W = 'embedded newline -- must not split a record'
           Value = '3'; Args = @('a1:D%="a\nb"') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
