# The hostile-string builders the fuzz cases share; dot-sourced by each
# fuzz .test.ps1 before its $case literal expands them.
# The hostile inputs, chosen for what they do to a DECODER, not for what a
# user would type. F is the formula; W is what the case is trying to break.
# REPT builds the long string AT CALC TIME. As a literal it would be vacuous:
# Excel refuses a string literal over 255 characters in a formula, so a case
# written that way never feeds 300 characters to the decoder at all.
$longAscii   = 'REPT("A",300)'              # past the 255 a byte count can hold
$longWide    = 'REPT("W",400)'              # long, but legal for a WCHAR count
# U+1F600 is beyond the BMP, so it is a SURROGATE PAIR in UTF-16 -- exactly
# the case a decoder that counts characters rather than code units gets wrong.
$unicode     = '"' + [char]0x00E9 + [char]0x4E2D + [char]0xD83D + [char]0xDE00 + 'x"'
$quoted      = '"he said ""hi"", ok"'        # commas and quotes: CSV escaping
$newline     = '"a' + [char]10 + 'b"'        # embedded newline
$empty       = '""'
$tiny        = '0.000000000000001'
$huge        = '1E307'
$negzero     = '-0'

