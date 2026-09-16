. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxD($longAscii)";    W = '300 chars into a 255-max byte count'
           # Excel refuses the call itself (#VALUE!), so there is no call to trace.
           Calls = @() }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
