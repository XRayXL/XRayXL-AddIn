$case = @{ Formula='=TxD("hello")'; Fn='TxD'; Value='5'; Args=@('a1:D="hello"'); Ret='5'
     Why='BYTE-COUNTED ASCII -- read as wide it fabricates a string' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
