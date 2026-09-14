$case = @{ Formula='=TxRetQNum(2)'; Fn='TxRetQNum'; Value='22'; Args=@('a1:B=2'); Ret='22'
     Why='XLOPER12 number return' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
