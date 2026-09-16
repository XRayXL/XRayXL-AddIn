$case = @{ Formula='=TxRefR(B2:C3)'; Fn='TxRefR'; Value='4'; Args=@('a1:R=SRef(R2C2:R3C3)'); Ret='4'; TypeText='R'; ArgCount='1'; RetType='Q'
     Why='the narrow R reference: a block on the calling sheet, whose XLREF has 16-bit rows and 8-bit columns' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
