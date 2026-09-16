$case = @{ Name='no-faults-under-load'
     Setup=@'
Public Sub T_Load(ByVal n As Long)
    Dim i As Long
    Dim s As Double
    For i = 1 To n
        s = s + Sqr(i)
    Next i
End Sub
'@
     Invoke=@{ Name='T_Load'; Args=@(200000) }
     Expect={ param($t)
        if ($t.faults -gt 0)    { return "$($t.faults) guarded trailer reads faulted under load" }
        if ($t.tableFull -gt 0) { return "procedure table filled" }
        if ($t.statements -lt 300000) { return "expected >=300k statements, got $($t.statements)" }
        $null }
     # Sqr is a VBA built-in, not a procedure: one activation.
     Calls=@( @{ Function='T_Load'; Args='a1:Long=200000'; Depth='1'; Parent=-1; Caller='none'; Outcome='returned' } )
     Why='the guarded read of rsp+0xB8 must hold for hundreds of thousands of
          statements without a single fault -- one fault means the offset is wrong' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
