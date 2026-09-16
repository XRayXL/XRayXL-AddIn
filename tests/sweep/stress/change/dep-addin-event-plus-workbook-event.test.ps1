$case = @{ Name='dep-addin-event-plus-workbook-event'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim z As Long
    z = 1
End Sub
'@
     Deps=@(
       @{ Name='AddFive'; IsAddin=$true; Modules=@{
         'MA'=@'
Public Function Unused(ByVal n As Double) As Double
    Unused = n
End Function
'@
       } }
     )
     Trigger=@{ Kind='Change'; Cell='B7'; Value=21 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change'; if ($why) { return $why }
        $null }
     # The add-in's Unused is never called: one handler, nothing else.
     Calls=@( @{ Function='Worksheet_Change'; Depth='1'; Parent=-1; Outcome='returned' } )
     Why='an add-in with a workbook-level event AND the host workbooks own handler' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
