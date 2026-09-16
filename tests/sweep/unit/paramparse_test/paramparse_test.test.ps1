# XRayXL_SetTraceParam's argument grammar: what is accepted, what is refused, and the BUFFERSIZE units.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'paramparse_test'
