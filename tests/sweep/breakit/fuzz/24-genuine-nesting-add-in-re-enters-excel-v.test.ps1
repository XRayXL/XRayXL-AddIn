. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCallsBack(5)";     W = 'GENUINE nesting: add-in re-enters Excel via xlUDF' ; NestDepth = 2
           # TxCallsBack(x) calls TxB(x,1) through xlUDF and returns its answer + 0.25.
           Calls = @(
               @{ Function = 'TxCallsBack'; Args = 'a1:B=5';        Ret = '51.25'; Depth = '1'; Parent = -1 }
               @{ Function = 'TxB';         Args = 'a1:B=5 a2:B=1'; Ret = '51';    Depth = '2'; Parent = 0 }
           ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
