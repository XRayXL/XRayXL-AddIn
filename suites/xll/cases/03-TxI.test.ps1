$case = @{ Formula='=TxI(7)'; Fn='TxI'; Value='7'; Args=@('a1:I=7'); Ret='7'
     Why='signed short, low 16 bits' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
