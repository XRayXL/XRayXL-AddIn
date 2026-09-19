$case = @{ Formula='=TxK({1,2;3,4})'; Fn='TxK'; Value='10'; Args=@('a1:K=Double[1..2,1..2]{{1,2},{3,4}}'); Ret='10'
     Why='FP -- 2-byte dimensions. Cross-decoded as FP12 it reports ZERO cells' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
