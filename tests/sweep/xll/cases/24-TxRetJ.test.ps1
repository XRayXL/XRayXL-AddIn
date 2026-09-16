$case = @{ Formula='=TxRetJ(2)'; Fn='TxRetJ'; Value='4'; Args=@('a1:B=2'); Ret='4'
     Why='int32 return, rax as an integer' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
