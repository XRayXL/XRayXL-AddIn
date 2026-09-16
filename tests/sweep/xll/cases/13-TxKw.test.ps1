$case = @{ Formula='=TxKw({1,2;3,4})'; Fn='TxKw'; Value='10'; Args=@('a1:K%=Double[1..2,1..2]{1,2,3,4}'); Ret='10'
     Why='FP12 -- 4-byte dimensions, and K% is TWO characters' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
