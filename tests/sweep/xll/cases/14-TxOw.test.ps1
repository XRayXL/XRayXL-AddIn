$case = @{ Formula='=TxOw({1,2;3,4})'; Fn='TxOw'; Value='10'; Args=@('a1:O%=Double[1..2,1..2]{{1,2},{3,4}}'); Ret='10'; TypeText='O%'; ArgCount='1'; RetType='Q'
     Why='ONE type code, THREE ABI slots. Miscounting shifts every later argument' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
