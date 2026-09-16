# VARIANT, SAFEARRAY and BSTR values render as trace text, nested arrays in full.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'decode_test'
