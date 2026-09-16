$case = @{ Name='event-worksheet-change'
     SheetCode=@'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim z As Long
    z = 1
    Helper
End Sub
Private Sub Helper()
    Dim q As Long
    q = 2
End Sub
'@
     Trigger=@{ Kind='Change'; Cell='B2'; Value=42 }
     Expect={ param($t)
        if ($t.procedures -lt 1) { return "no procedure traced from the Change event" }
        $n = $t.rows | Where-Object { $_.function -eq 'Worksheet_Change' }
        if (-not $n) { return "Worksheet_Change not named in the trace" }
        $null }
     Calls=@(
        @{ Function='Worksheet_Change'; Depth='1'; Parent=-1; Outcome='returned' }
        @{ Function='Helper';           Depth='2'; Parent=0;  Outcome='returned' } )
     Why='the plainest event path: Worksheet_Change, which no Application.Run reaches' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
