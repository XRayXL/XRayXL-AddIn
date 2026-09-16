# One saved workbook, XLL and VBA UDFs together, single-threaded calculation:
# one file, one monotonic sequence, both sources, named and attributed.
. (Join-Path $PSScriptRoot '_driver.ps1')
Invoke-TimelineTest
