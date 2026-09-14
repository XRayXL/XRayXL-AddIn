$case = @{ Formula='=TxRetE(2)'; Fn='TxRetE'; Value='10'; Args=@('a1:B=2'); Ret='10'
     Why='rax as a pointer to double' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
