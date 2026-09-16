# The trace row formatter: column count, order, CSV escaping and the header.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'rowcsv_test'
