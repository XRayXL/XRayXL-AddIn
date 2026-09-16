. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxCallsBack2(5)";    W = 'TWO levels of re-entry' ; NestDepth = 3
           # TxCallsBack2(x) doubles TxCallsBack(x), which is TxB(x,1) + 0.25.
           Calls = @(
               @{ Function = 'TxCallsBack2'; Args = 'a1:B=5';        Ret = '102.5'; Depth = '1'; Parent = -1 }
               @{ Function = 'TxCallsBack';  Args = 'a1:B=5';        Ret = '51.25'; Depth = '2'; Parent = 0 }
               @{ Function = 'TxB';          Args = 'a1:B=5 a2:B=1'; Ret = '51';    Depth = '3'; Parent = 1 }
           ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
