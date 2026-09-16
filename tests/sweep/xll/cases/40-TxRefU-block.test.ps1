$case = @{ Formula='=TxRefU(B2:C3)'; Fn='TxRefU'; Value='4'; Args=@('a1:U=SRef(R2C2:R3C3)'); Ret='4'; TypeText='U'; ArgCount='1'; RetType='Q'
     Why='a U argument receives the reference, not its value: a block on the calling sheet, counted by the add-in from the structure Excel passed' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
