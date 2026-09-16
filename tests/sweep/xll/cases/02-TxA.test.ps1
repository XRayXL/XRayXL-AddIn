$case = @{ Formula='=TxA(TRUE)'; Fn='TxA'; Value='1'; Args=@('a1:A=1'); Ret='1'
     Why='boolean as a short in an integer register' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
