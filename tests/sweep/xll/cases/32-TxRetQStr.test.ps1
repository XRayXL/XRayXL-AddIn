$case = @{ Formula='=TxRetQStr(2)'; Fn='TxRetQStr'; Value='q:2'; Args=@('a1:B=2'); Ret='"q:2"'
     Why='XLOPER12 string return -- WCHAR-counted' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
