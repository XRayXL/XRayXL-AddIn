# Hostile inputs the fuzz cases share, chosen for what they do to a decoder, not what a user types.
# F is a case's formula; W is what it is trying to break.
# REPT builds the long strings at calc time: Excel refuses a formula string literal over 255
# characters, so a literal would never feed 300 of them to the decoder.
$longAscii   = 'REPT("A",300)'              # past the 255 a byte count can hold
$longWide    = 'REPT("W",400)'              # long, but legal for a WCHAR count
# U+1F600 is a surrogate pair in UTF-16: a decoder counting characters rather than code units gets it wrong.
$unicode     = '"' + [char]0x00E9 + [char]0x4E2D + [char]0xD83D + [char]0xDE00 + 'x"'
$quoted      = '"he said ""hi"", ok"'        # commas and quotes: CSV escaping
$newline     = '"a' + [char]10 + 'b"'        # embedded newline
$empty       = '""'
$tiny        = '0.000000000000001'
$huge        = '1E307'
$negzero     = '-0'

