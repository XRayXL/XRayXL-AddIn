. (Join-Path $PSScriptRoot '..\_strings.ps1')
$case = @{ F = "=TxMany(1,2,3,4,5,6,7,8,9,10,11,12)"; W = 'TWELVE doubles: eight on the stack'
           # Position-weighted: the sum of i*i for i = 1..12 is 650.
           Calls = @( @{ Function = 'TxMany'; Ret = '650'; Depth = '1'; Parent = -1
                         Args = 'a1:B=1 a2:B=2 a3:B=3 a4:B=4 a5:B=5 a6:B=6 a7:B=7 a8:B=8 a9:B=9 a10:B=10 a11:B=11 a12:B=12' } ) }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
