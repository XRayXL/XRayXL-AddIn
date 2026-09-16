$case = @{ Formula='=TxRetQErr(0)'; Fn='TxRetQErr'; Value='#N/A'; Args=@('a1:B=0'); Ret='#N/A'
     Why='an ERROR is an outcome, and must be reported as one' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
