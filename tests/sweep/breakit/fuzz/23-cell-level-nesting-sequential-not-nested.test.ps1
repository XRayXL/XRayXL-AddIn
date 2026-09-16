. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxB(TxB(1,2),3)";    W = 'cell-level nesting (sequential, not nested)' ; NestDepth = 1
           # Excel finishes the inner call before the outer starts: two calls, both at depth 1.
           Calls = @(
               @{ Function = 'TxB'; Args = 'a1:B=1 a2:B=2';  Ret = '12';  Depth = '1'; Parent = -1 }
               @{ Function = 'TxB'; Args = 'a1:B=12 a2:B=3'; Ret = '123'; Depth = '1'; Parent = -1 }
           ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
