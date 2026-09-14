$case = @{ Name='dep-cross-workbook-run'
     Modules=@{
       'M'=@'
Public Sub LocalGo()
    Application.Run "DepD.xlsm!RemoteMac"
End Sub
'@
     }
     Deps=@(
       @{ Name='DepD'; Modules=@{
         'MD'=@'
Public Sub RemoteMac()
    Dim z As Long
    z = 1
    RemoteHelper
End Sub
Public Sub RemoteHelper()
    Dim q As Long
    q = 2
End Sub
'@
       } }
     )
     Trigger=@{ Kind='Run'; Name='LocalGo' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'LocalGo','RemoteMac'; if ($why) { return $why }
        $e = $t.rows | Where-Object { $_.function -eq 'RemoteMac' } | Select-Object -First 1
        if ($e.module -notmatch 'DepD') { return "remote macro attributed to [$($e.module)]" }
        $null }
     Why='a macro in A calling Application.Run against a macro in B' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
