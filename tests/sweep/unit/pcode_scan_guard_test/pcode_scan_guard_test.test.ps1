# The arm-time p-code scan reads handlers safely and declines an unreadable page.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'pcode_scan_guard_test'
