$case = @{ Formula='=TxQ("text")'; Fn='TxQ'; Value='2'; Args=@('a1:Q="text"'); Ret='2'
     Why='XLOPER12 carrying a string' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
