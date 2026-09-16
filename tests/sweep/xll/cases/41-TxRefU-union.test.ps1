$case = @{ Formula='=TxRefU((B2:C3,E5))'; Fn='TxRefU'; Value='5'; Args=@('a1:U=Ref(R2C2:R3C3,R5C5:R5C5)'); Ret='5'; TypeText='U'; ArgCount='1'; RetType='Q'
     Why='a U argument given a union of two areas: the add-in counts every area of the reference Excel passed' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
