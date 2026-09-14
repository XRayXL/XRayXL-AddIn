$case = @{ Formula='=TxJ(70000)'; Fn='TxJ'; Value='70000'; Args=@('a1:J=70000'); Ret='70000'
     Why='signed int32, low 32 bits -- a value that does NOT fit in 16' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
