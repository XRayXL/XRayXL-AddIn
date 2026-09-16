$case = @{ Formula='=TxCw("hello")'; Fn='TxCw'; Value='5'; Args=@('a1:C%="hello"'); Ret='5'
     Why='null-terminated UTF-16, and C% is TWO characters of type text' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
