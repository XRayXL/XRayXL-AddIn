# The walk checks the last statement ends at ProcSize, and names the exit when it does not.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'pcode_closure_test'
