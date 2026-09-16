$case = @{ Formula='=TxQ(TRUE)'; Fn='TxQ'; Value='3'; Args=@('a1:Q=TRUE'); Ret='3'
     Why='XLOPER12 carrying a boolean' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
