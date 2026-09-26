# A parameter only passed on is typed by the Variant label after its push; a typed load still decides.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'pcode_label_test'
