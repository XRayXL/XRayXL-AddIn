# Every value xlfCaller can return, decoded as the product decodes it.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'callerdecode'
