# The Perfetto page the ribbon writes: the shipping page with a trace's exact bytes inside it,
# and refusals that leave nothing behind. No Excel, no browser.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'perfettopage_test'
