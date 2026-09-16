$case = @{ Formula='=TxDw("hello")'; Fn='TxDw'; Value='5'; Args=@('a1:D%="hello"'); Ret='5'
     Why='WCHAR-counted UTF-16 -- what Excel-DNA registers strings as' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
