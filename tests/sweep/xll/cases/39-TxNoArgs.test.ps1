$case = @{ Formula='=TxNoArgs()'; Fn='TxNoArgs'; Value='42'; Args=@(); Ret='42'; TypeText=''; ArgCount='0'; RetType='Q'
     Why='no arguments: an empty typetext and argcount 0, the same as a VBA procedure with no parameters' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
