$case = @{ Formula='=TxHiddenName(1)'; Fn='TxHiddenName'; Value='4242'; Args=@('a1:A=1'); Ret='4242'
     Why='DISPLAY NAME DIFFERS FROM THE EXPORT. TracedAddin exports TxUnregistered and registers it as TxHiddenName, so the trace must report the name the user typed and not the export. Every other case here registers under its own export name, which means they all pass whether the name is looked up or merely echoed -- this is the only case that can tell those apart, and it is the Excel-DNA shape (f0 carrying a real name)' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-XllCase $case
