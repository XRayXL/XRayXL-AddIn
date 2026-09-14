. (Join-Path $PSScriptRoot '..\_strings.ps1')
# Its NUL is the last byte read, so the string is whole and must not be marked cut.
$case = @{ F = '=TxC(REPT("a",127))'; W = '127 chars, the longest that fits the 128-byte read with its NUL'
           ArgsExact = 'a1:C="aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
