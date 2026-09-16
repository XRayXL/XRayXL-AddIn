$case = @{ Formula='=TxE(3.5)'; Fn='TxE'; Value='3.5'; Args=@('a1:E=3.5'); Ret='3.5'
     Why='pointer to double, NOT a double in XMM' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
