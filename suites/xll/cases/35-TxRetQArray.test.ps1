$case = @{ Formula='=TxRetQArray(1)'; Fn='TxRetQArray'; Value='1'; Args=@('a1:B=1'); Ret='Variant[1..2,1..2]{'
     Why='array CONTENTS are readable at the add-in boundary, not just the shape' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
