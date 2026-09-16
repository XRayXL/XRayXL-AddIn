# Closing the trace session while writer threads keep depositing rows neither faults nor hangs.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'csv_close_race_test'
