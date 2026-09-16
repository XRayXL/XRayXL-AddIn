. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxOptional(5)";      W = 'omitted trailing argument'
           # An omitted argument arrives as Missing, and TxOptional then returns the first alone.
           Calls = @( @{ Function = 'TxOptional'; Args = 'a1:Q=5 a2:Q=Missing'; Ret = '5'; Depth = '1'; Parent = -1 } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
