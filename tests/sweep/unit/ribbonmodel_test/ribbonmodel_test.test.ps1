# The ribbon's decisions -- enablement, what each control reads and writes, and
# the cross-check between the shipped customUI XML and the handlers. No Excel.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'ribbonmodel_test'
