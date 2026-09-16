$case = @{ Formula='=TxOptional(5)'; Fn='TxOptional'; Value='5'; Args=@('a1:Q=5','a2:Q=Missing'); Ret='5'
     Why='an omitted argument must read as missing, not as zero' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
