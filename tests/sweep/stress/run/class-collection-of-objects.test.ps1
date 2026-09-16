$case = @{ Name='class-collection-of-objects'
     Modules=@{
       'M'=@'
Public Sub Go()
    Dim c As Collection
    Dim it As CItem
    Dim i As Long
    Dim total As Long
    Set c = New Collection
    For i = 1 To 200
        Set it = New CItem
        it.SetId i
        c.Add it
    Next i
    For i = 1 To 200
        total = total + c.Item(i).GetId()
    Next i
End Sub
'@
     }
     Classes=@{
       'CItem'=@'
Private mId As Long
Private Sub Class_Initialize()
    mId = 0
End Sub
Public Sub SetId(ByVal i As Long)
    mId = i
End Sub
Public Function GetId() As Long
    GetId = mId
End Function
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 3) { return "expected >=3 procedures, got $($t.procedures)" }
        if ($t.statements -lt 600) { return "expected >=600 statements, got $($t.statements)" }
        $null }
     # Each New runs Class_Initialize inside Go, then SetId; afterwards GetId reads each id back.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(1..200 | ForEach-Object {
               @{ Function='Class_Initialize'; Depth='2'; Parent=0; Outcome='returned' }
               @{ Function='SetId'; Args="a1:Long=$_"; Depth='2'; Parent=0; Outcome='returned' } }) +
           @(1..200 | ForEach-Object { @{ Function='GetId'; Ret="$_"; RetType='Long'; Depth='2'; Parent=0; Outcome='returned' } })
     Why='200 class instances held in a Collection, each constructed and used' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
