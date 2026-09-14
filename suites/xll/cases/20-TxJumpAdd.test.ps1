# Exported through a jump table (jumptable.asm): the export is a jump, and
# incremental linking does the same to direct exports, so hooks must follow it.
$case = @{ Formula='=TxJumpAdd(2,3)'; Fn='TxJumpAdd'; Value='5.5'; Args=@('a1:B=2','a2:B=3'); Ret='5.5'
     Why='an export that is a jump to the real function must hook like any other' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
