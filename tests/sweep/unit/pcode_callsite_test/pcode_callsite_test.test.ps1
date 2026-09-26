# A caller's call site: the pushes before the call, a local's type, and each literal's type.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'pcode_callsite_test'
