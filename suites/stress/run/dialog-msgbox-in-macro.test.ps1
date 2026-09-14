$case = @{ Name='dialog-msgbox-in-macro'
     Modules=@{
       'M'=@'
Public Sub ShowBox()
    Dim z As Long
    z = 1
    MsgBox "stress harness: this box is dismissed automatically"
    z = 2
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='ShowBox'; MayRaise=$true; ExpectDialog=$true }
     Expect={ param($t)
        if ($t.dialogs -lt 1) { return "no dialog recorded" }
        $why = Assert-VbaTraced $t 'ShowBox'; if ($why) { return $why }
        $null }
     Why='a macro that calls MsgBox: Excel blocks mid-frame until the box is dismissed' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
