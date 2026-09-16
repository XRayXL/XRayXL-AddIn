. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=SUM(TxB(1,2),TxB(3,4))"; W = 'two traced calls inside one built-in' ; NestDepth = 1
           # TxB returns a*10 + b; SUM evaluates both, left to right, neither inside the other.
           Calls = @(
               @{ Function = 'TxB'; Args = 'a1:B=1 a2:B=2'; Ret = '12'; Depth = '1'; Parent = -1; Caller = 'cell'; Cell = 'A1' }
               @{ Function = 'TxB'; Args = 'a1:B=3 a2:B=4'; Ret = '34'; Depth = '1'; Parent = -1; Caller = 'cell'; Cell = 'A1' }
           ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
