# A double with more digits than `%g` prints. Both columns land in the same trace file, so the
# XLL and VBA renderings of 1234567.89012345 must agree; `%g` would keep six of its fifteen
# significant digits. TxB returns a*10 + b, so one formula checks both the argument and the
# return.
$case = @{ Formula='=TxB(1234567.89012345,0)'; Fn='TxB'
     # the cell's value to fifteen significant digits, not the display text
     Value='12345678.9012345'
     Args=@('a1:B=1234567.89012345','a2:B=0')
     Ret='12345678.9012345'
     Why='a double with fifteen significant digits -- %g keeps six, and the VBA
          column keeps all of them, so the two columns of one trace file would
          otherwise disagree about the same number' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
