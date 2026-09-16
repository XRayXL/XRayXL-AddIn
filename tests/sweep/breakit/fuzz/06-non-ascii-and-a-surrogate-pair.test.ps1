. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCw($unicode)";     W = 'non-ASCII and a surrogate pair'
           Value = '5'; Args = @('a1:C%="\u00E9\u4E2D\uD83D\uDE00x"')
           # wcslen counts code units: the surrogate pair is two. Non-ASCII is written as \u escapes.
           Calls = @( @{ Function = 'TxCw'; Args = 'a1:C%="\u00E9\u4E2D\uD83D\uDE00x"'; Ret = '5'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
