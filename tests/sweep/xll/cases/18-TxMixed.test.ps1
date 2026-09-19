# D% is length-prefixed, so the first element is 3, not 'a': (1+2+3)*2 + 3 = 15
$case = @{ Formula='=TxMixed("abc",{1,2,3},2)'; Fn='TxMixed'; Value='15'
     Args=@('a1:D%="abc"','a2:K%=Double[1..1,1..3]{{','a3:E=2'); Ret='15'; TypeText='D%,K%,E'; ArgCount='3'; RetType='Q'
     Why='THE case: QD%K%E reads as FIVE arguments if parsed per character' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
