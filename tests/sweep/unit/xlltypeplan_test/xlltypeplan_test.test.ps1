# A '%' widens only the registration codes the C API gives a wide form.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'xlltypeplan_test'
