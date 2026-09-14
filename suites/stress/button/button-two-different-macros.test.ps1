$case = @{ Name='button-two-different-macros'
     Modules=@{
       'M'=@'
Public Sub BtnA()
    Dim z As Long
    z = 1
End Sub
Public Sub BtnB()
    Dim z As Long
    z = 2
End Sub
'@
     }
     Shapes=@(@{ Index=0; Name='B1'; OnAction='BtnA' }, @{ Index=1; Name='B2'; OnAction='BtnB' })
     Trigger=@{ Kind='Button'; Name='B1' }
     Expect={ param($t)
        $names = @(Get-VbaEntryNames $t)
        $why = Assert-VbaTraced $t 'BtnA'; if ($why) { return $why }
        if ($names -contains 'BtnB')    { return "the OTHER button ran, which it must not" }
        $null }
     Why='two buttons on one sheet with different macros: only the pressed one runs' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
