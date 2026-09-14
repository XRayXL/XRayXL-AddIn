$case = @{ Formula='=TxStackArgs(1,2,3,4,5)'; Fn='TxStackArgs'; Value='12345'
     Args=@('a1:B=1','a2:B=2','a3:B=3','a4:B=4','a5:B=5'); Ret='12345'
     Why='the FIFTH argument is on the caller stack, not in a register -- it is
          also what the thunk copy loop exists for. (Named TxStackArgs, not
          TxB5: Excel refuses to register a name that is a valid cell
          reference, and column TXB is inside XFD. rc=32, xlretFailed.)' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
