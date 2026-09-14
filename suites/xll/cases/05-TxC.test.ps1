$case = @{ Formula='=TxC("hello")'; Fn='TxC'; Value='5'; Args=@('a1:C="hello"'); Ret='5'
     Why='null-terminated ASCII' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
