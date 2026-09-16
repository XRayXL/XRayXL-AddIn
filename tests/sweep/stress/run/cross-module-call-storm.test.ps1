$case = @{ Name='cross-module-call-storm'
     Modules=@{
       'MA'=@'
Public Sub FA(ByVal n As Long)
    If n > 0 Then MB.FB n - 1
End Sub
'@
       'MB'=@'
Public Sub FB(ByVal n As Long)
    If n > 0 Then MC.FC n - 1
End Sub
'@
       'MC'=@'
Public Sub FC(ByVal n As Long)
    If n > 0 Then MD.FD n - 1
End Sub
'@
       'MD'=@'
Public Sub FD(ByVal n As Long)
    If n > 0 Then MA.FA n - 1
End Sub
'@
       'M'=@'
Public Sub Go()
    MA.FA 250
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 5) { return "expected >=5 procedures, got $($t.procedures)" }
        if ($t.unnamed -gt 0) { return "$($t.unnamed) procedure(s) could not be named under churn" }
        $null }
     # Go, then FA 250, FB 249, FC 248, FD 247, FA 246 ... down to 0: 251 activations, each inside the last.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(0..250 | ForEach-Object { @{ Function=@('FA','FB','FC','FD')[$_ % 4]; Args="a1:Long=$(250 - $_)"; Depth="$($_ + 2)"; Parent=$_; Outcome='returned' } })
     Why='four modules calling round-robin 250 times: constant module-identity churn' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
