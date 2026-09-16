# The procedure table: dedup, distinct slots, collisions, capacity and Reset.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'proctable_test'
