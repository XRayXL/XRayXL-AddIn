$case = @{ Formula='=TxM(5)'; Fn='TxM'; Value='5'; Args=@('a1:M=5'); Ret='5'
     Why='pointer to short' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
