$case = @{ Name='event-activex-checkbox'
     SheetCode=@'
Private Sub CB1_Click()
    Dim z As Long
    z = 1
    AfterClick
End Sub
Private Sub AfterClick()
    Dim q As Long
    q = 2
End Sub
'@
     ActiveX='CB1'
     Trigger=@{ Kind='ActiveX'; Name='CB1' }
     Expect={ param($t)
        $names = @(Get-VbaEntryNames $t)
        if (-not ($names -match 'CB1_')) { return "no ActiveX handler traced; saw [$($names -join ',')]" }
        $null }
     Why='an ActiveX control event sink -- a different dispatch from Application.Run' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
