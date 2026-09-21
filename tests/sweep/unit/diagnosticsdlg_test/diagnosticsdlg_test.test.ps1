# The Diagnostics dialog, opened and driven outside Excel: the three pages, the search, the
# column sort, the detail line and Refresh; and the module, environment and process tables and
# their CSV, JSON and clipboard forms on their own.
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-UnitTest 'diagnosticsdlg_test'
