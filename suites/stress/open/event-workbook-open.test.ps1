$case = @{ Name='event-workbook-open'
     ThisWbCode=@'
Private Sub Workbook_Open()
    Dim z As Long
    z = 1
    OpenHelper
End Sub
Private Sub OpenHelper()
    Dim q As Long
    q = 2
End Sub
'@
     Trigger=@{ Kind='Open' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Workbook_Open'; if ($why) { return $why }
        $null }
     Why='Workbook_Open, which fires before anything else can be armed by hand' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
