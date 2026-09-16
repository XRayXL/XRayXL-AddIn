# A DOUBLE WITH MORE DIGITS THAN `%g` PRINTS -- the XLL column against the VBA one.
#
# Both columns land in the SAME trace file, so a reader comparing an XLL row
# with a VBA row is comparing two renderings that must mean the same thing. The
# VBA side was measured and fixed: an argument printed at `%g` -- six
# significant figures -- while the identical value as a result printed at
# `%.15g`, so 1234567.89012345 arrived as `1.23457e+06`.
#
# `src/xll/xlldecode.cpp` was not in that change and still formats with `%g`. This is
# the case that says whether that matters in practice: 1234567.89012345 has
# fifteen significant digits, and `%g` keeps six.
#
# THE LITERAL IS CHOSEN TO SEPARATE THE TWO FORMATS. A case built on 2 or 3 --
# which every other TxB case uses -- passes against either, which is exactly why
# this went unnoticed on the VBA side until a value with enough digits was
# planted.
#
# TxB returns a*10 + b, so the RESULT carries even more digits than the
# argument: one formula measures both the argument and the return.
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
