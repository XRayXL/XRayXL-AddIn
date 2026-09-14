$case = @{ Formula='=TxRetD(2)'; Fn='TxRetD'; Value='d:2'; Args=@('a1:B=2'); Ret='"d:2"'
     Why='rax as a pointer to byte-counted ASCII' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
