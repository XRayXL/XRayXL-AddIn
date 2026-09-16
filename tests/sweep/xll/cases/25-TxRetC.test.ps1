$case = @{ Formula='=TxRetC(2)'; Fn='TxRetC'; Value='c:2'; Args=@('a1:B=2'); Ret='"c:2"'
     Why='rax as a RAW POINTER to ASCII -- the row the old guard passed and faked' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
