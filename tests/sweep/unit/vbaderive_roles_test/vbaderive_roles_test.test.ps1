# No p-code handler is patched in two roles.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'vbaderive_roles_test'
