$case = @{ Formula='=TxRetB(2)'; Fn='TxRetB'; Value='6'; Args=@('a1:B=2'); Ret='6'
     Why='a double return lives in XMM0. Reading rax invents a value' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
