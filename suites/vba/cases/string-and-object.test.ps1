$case = @{ Name='string-and-object'
     Setup=@'
Public Sub T_Mixed()
    Dim s As String
    Dim c As Object
    Dim i As Long
    s = "abc" & "def"
    Set c = ThisWorkbook.Worksheets(1)
    For i = 1 To 50
        s = s & "x"
    Next i
    Set c = Nothing
End Sub
'@
     Invoke=@{ Name='T_Mixed'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) faults with strings and objects in play" }
        if ($t.statements -lt 100) { return "expected >=100 statements, got $($t.statements)" }
        $null }
     Why='strings and COM objects take different interpreter paths from arithmetic' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
