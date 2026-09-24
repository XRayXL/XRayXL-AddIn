# The same exit criterion under multithreaded calculation, where threads reach the writer's lock
# in another order and the file must still be one trace. Refuses to pass if MTC never engaged.
. (Join-Path $PSScriptRoot '_driver.ps1')
Invoke-TimelineTest -Mtc
