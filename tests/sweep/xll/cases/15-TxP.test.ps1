$case = @{ Formula='=TxP(1.5)'; Fn='TxP'; Value='1'; Args=@('a1:P=1.5'); Ret='1'
     Why='XLOPER -- 24 bytes, xltype at +16. A DIFFERENT struct from XLOPER12' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
