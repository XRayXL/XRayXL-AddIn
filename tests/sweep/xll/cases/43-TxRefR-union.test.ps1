$case = @{ Formula='=TxRefR((B2:C3,E5))'; Fn='TxRefR'; Value='5'; Args=@('a1:R=Ref(R2C2:R3C3,R5C5:R5C5)'); Ret='5'; TypeText='R'; ArgCount='1'; RetType='Q'
     Why='the narrow R reference given a union of two areas, read through XLMREF' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
