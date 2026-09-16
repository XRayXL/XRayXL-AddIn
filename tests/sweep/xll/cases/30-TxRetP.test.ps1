$case = @{ Formula='=TxRetP(2)'; Fn='TxRetP'; Value='14'; Args=@('a1:B=2'); Ret='14'
     Why='a NARROW XLOPER return. Reading +24 is past the end of a 24-byte struct' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
