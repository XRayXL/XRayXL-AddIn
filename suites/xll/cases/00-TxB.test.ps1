$case = @{ Formula='=TxB(2,3)'; Fn='TxB'; Value='23'; Args=@('a1:B=2','a2:B=3'); Ret='23'; TypeText='B,B'; ArgCount='2'; RetType='Q'
     Why='two doubles in XMM by POSITION -- the ordering rule' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
