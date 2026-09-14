$case = @{ Formula='=TxJumpStr("hello")'; Fn='TxJumpStr'; Value='500'; Args=@('a1:D%="hello"'); Ret='500'
     Why='same, with a counted-string argument decoded through the thunk' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
