$case = @{ Name='button-deep-chain'
     Modules=@{
       'M'=@'
Public Sub D00()
    D01
End Sub
Public Sub D01()
    D02
End Sub
Public Sub D02()
    D03
End Sub
Public Sub D03()
    D04
End Sub
Public Sub D04()
    D05
End Sub
Public Sub D05()
    D06
End Sub
Public Sub D06()
    D07
End Sub
Public Sub D07()
    D08
End Sub
Public Sub D08()
    D09
End Sub
Public Sub D09()
    D10
End Sub
Public Sub D10()
    Dim z As Long
    z = 1
End Sub
'@
     }
     Shapes=@(@{ Index=0; Name='B1'; OnAction='D00' })
     Trigger=@{ Kind='Button'; Name='B1' }
     Expect={ param($t)
        if ($t.maxDepth -lt 10) { return "expected depth >=10, got $($t.maxDepth)" }
        $null }
     Why='a button macro that runs a ten-deep call chain' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
