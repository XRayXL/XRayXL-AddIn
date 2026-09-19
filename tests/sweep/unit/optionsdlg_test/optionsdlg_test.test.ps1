# The Options dialog, opened and driven outside Excel: its pages, the Format drop-down, the
# embedded licence and notices, OK, Cancel and the armed lock; and the reflow of their text.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'optionsdlg_test'
