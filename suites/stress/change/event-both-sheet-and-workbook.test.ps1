$case = @{ Name='event-both-sheet-and-workbook'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim z As Long
    z = 1
End Sub
'@
     ThisWbCode=@'
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim q As Long
    q = 2
End Sub
'@
     Trigger=@{ Kind='Change'; Cell='B4'; Value=12 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Worksheet_Change','Workbook_SheetChange'; if ($why) { return $why }
        $null }
     Why='one edit firing BOTH the sheet handler and the workbook handler' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
