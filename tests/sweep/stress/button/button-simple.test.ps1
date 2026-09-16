$case = @{ Name='button-simple'
     Modules=@{
       'M'=@'
Public Sub BtnOne()
    Dim z As Long
    z = 1
    BtnHelper
End Sub
Public Sub BtnHelper()
    Dim q As Long
    q = 2
End Sub
'@
     }
     Shapes=@(@{ Index=0; Name='B1'; OnAction='BtnOne' })
     Trigger=@{ Kind='Button'; Name='B1' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'BtnOne'; if ($why) { return $why }
        $null }
     Calls=@(
        @{ Function='BtnOne';    Depth='1'; Parent=-1; Outcome='returned' }
        @{ Function='BtnHelper'; Depth='2'; Parent=0;  Outcome='returned' } )
     Why='a Form button macro through Excels macro dispatch, where a hit-tested click lands' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
