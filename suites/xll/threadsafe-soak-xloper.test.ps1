# SOAK: the same with XLOPER12 (Q) arguments and result, unpacked and allocated per call on worker
# threads. ROW()-4 and COLUMN()-4 make the edge cells answer with a message naming the bad argument.
. (Join-Path $PSScriptRoot '_soak.ps1')
Invoke-XllSoak -Fn 'TxRowColQ' -ArgFormat 'a1:Q={0} a2:Q={1}' -TypeText 'Q,Q' -RetType 'Q$' -Offset 4 -Natural
