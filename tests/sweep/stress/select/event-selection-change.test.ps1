$case = @{ Name='event-selection-change'
     SheetCode=@'
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim z As Long
    z = Target.Row
End Sub
'@
     Trigger=@{ Kind='Select'; Cell='D9' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_SelectionChange'; if ($why) { return $why }
        $null }
     Calls=@( @{ Function='Worksheet_SelectionChange'; Depth='1'; Parent=-1; Outcome='returned' } )
     Why='Worksheet_SelectionChange: a trigger with no cell edit behind it at all' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
