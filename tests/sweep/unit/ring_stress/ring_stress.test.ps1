# The lock-free output ring under many producers: nothing torn, lost, duplicated or miscounted.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'ring_stress'
