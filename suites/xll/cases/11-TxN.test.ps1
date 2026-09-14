$case = @{ Formula='=TxN(9)'; Fn='TxN'; Value='9'; Args=@('a1:N=9'); Ret='9'
     Why='pointer to int32' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
