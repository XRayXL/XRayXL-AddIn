# SOAK: a thread-safe XLL function taking two doubles (B), 20,000 cells x 20 passes on 8 threads.
. (Join-Path $PSScriptRoot '_soak.ps1')
Invoke-XllSoak -Fn 'TxRowCol' -ArgFormat 'a1:B={0} a2:B={1}' -TypeText 'B,B' -RetType 'Q$'
