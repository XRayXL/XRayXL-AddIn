# A NEGATIVE DOUBLE SMALLER THAN ONE, as an argument and as the result.
#
# Every other TxB case plants a positive value, so a decoder that lost the sign of
# a fraction would pass all of them. TxB returns a*10 + b: -0.05*10 + 0.3 = -0.2.
$case = @{ Formula='=TxB(-0.05,0.3)'; Fn='TxB'; Value='-0.2'
     Args=@('a1:B=-0.05','a2:B=0.3'); Ret='-0.2'
     Why='a negative fraction keeps its sign -- -0.05 and 0.05 differ only there' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
