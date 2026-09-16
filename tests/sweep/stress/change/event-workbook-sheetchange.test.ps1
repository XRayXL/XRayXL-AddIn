$case = @{ Name='event-workbook-sheetchange'
     ThisWbCode=@'
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim z As Long
    z = 1
End Sub
'@
     Trigger=@{ Kind='Change'; Cell='B3'; Value=11 }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Workbook_SheetChange'; if ($why) { return $why }
        $null }
     Calls=@( @{ Function='Workbook_SheetChange'; Depth='1'; Parent=-1; Outcome='returned' } )
     Why='the WORKBOOK-level SheetChange, a different sink from the sheet-level one' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
