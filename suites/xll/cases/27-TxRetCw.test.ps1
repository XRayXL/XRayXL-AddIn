$case = @{ Formula='=TxRetCw(2)'; Fn='TxRetCw'; Value='cw:2'; Args=@('a1:B=2'); Ret='"cw:2"'
     Why='rax as a pointer to null-terminated UTF-16' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
