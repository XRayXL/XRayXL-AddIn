$case = @{ Formula='=TxRetDw(2)'; Fn='TxRetDw'; Value='dw:2'; Args=@('a1:B=2'); Ret='"dw:2"'
     Why='rax as a pointer to WCHAR-counted UTF-16' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
