# The same exit criterion under MULTITHREADED calculation -- a different test,
# not a repeat: several threads stamp concurrently and reach the writer's lock
# in another order, and what must still hold is everything that makes the file
# one trace. Refuses to pass if MTC never actually engaged.
. (Join-Path $PSScriptRoot '_driver.ps1')
Invoke-TimelineTest -Mtc
