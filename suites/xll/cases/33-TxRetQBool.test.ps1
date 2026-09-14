$case = @{ Formula='=TxRetQBool(1)'; Fn='TxRetQBool'; Value='TRUE'; Args=@('a1:B=1'); Ret='TRUE'
     Why='XLOPER12 boolean return' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
