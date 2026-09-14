. (Join-Path $PSScriptRoot '..\_strings.ps1')
# wcslen counts UTF-16 units, so the surrogate pair makes it 5; each unit above 126 renders as \uNNNN.
$case = @{ F = "=TxCw($unicode)";     W = 'non-ASCII and a surrogate pair'
           Value = '5'; Args = @('a1:C%="\u00E9\u4E2D\uD83D\uDE00x"') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
