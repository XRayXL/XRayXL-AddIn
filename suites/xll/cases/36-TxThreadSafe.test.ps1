$case = @{ Formula='=TxThreadSafe(2)'; Fn='TxThreadSafe'; Value='4'; Args=@('a1:B=2'); Ret='4'; TypeText='B'; ArgCount='1'; RetType='Q$'
     Why='registered thread-safe, so Excel may run it on a worker thread' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
