$case = @{ Formula='=TxRetI(2)'; Fn='TxRetI'; Value='3'; Args=@('a1:B=2'); Ret='3'
     Why='short return, rax as an integer' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
