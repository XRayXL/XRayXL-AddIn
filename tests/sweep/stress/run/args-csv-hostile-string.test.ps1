$case = @{ Name='args-csv-hostile-string'
     Modules=@{
       'M'=@'
Public Sub Hostile(ByVal s As String)
    Dim z As Long
    z = Len(s)
End Sub
Public Sub Go()
    Hostile "a,b" & Chr(34) & "quoted" & Chr(34) & ",c" & vbCrLf & "second,line" _
          & Chr(9) & Chr(1) & "C:\temp" & ChrW(233) & ChrW(8364)
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Hostile' } | Select-Object -First 1
        if (-not $e) { return "no row for Hostile -- the CSV may be unparseable" }
        if (-not $e.proc) { return "proc column lost -- escaping broke the row" }
        if (-not $e.args) { return "args column lost -- the slot was not even captured" }

        # A broken escape fails in Read-TraceFile, which checks every line's field count against
        # the header, before Expect is reached.
        #
        # The value itself is asserted too: a declared String holding vbCrLf must decode as
        # text, not fall back to the raw qword. Control characters are rendered, not passed
        # through: the CSV writer would turn a raw CR or LF into a space.
        $want = 'a1:String="a,b\"quoted\",c\r\nsecond,line\t\x01C:\\temp\u00E9\u20AC"'
        if ($e.args -ne $want) {
            return "args decoded as '$($e.args)', expected '$want'" }
        $null }
     Why='a runtime-built string full of commas, quotes and newlines must not
          break the tracer or the shape of its row' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
