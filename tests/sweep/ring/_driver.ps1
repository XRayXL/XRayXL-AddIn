# The ring cases' shared scaffold: size the ring, pick a policy, run one VBA driver, and reconcile
# Disarm, the log and the file. The underscore keeps this out of test discovery.

# One argument-less activation per call: lean rows that any sane ring holds.
$RingLeanModule = @'
Public Sub Leaf()
    Dim t As Long
    t = 1
End Sub

Public Sub Drive()
    Dim i As Long
    For i = 1 To 800
        Leaf
    Next i
End Sub
'@

# A 2000-element array argument renders to ~9 KB, so one row nearly fills a 16 KB ring.
$RingFatModule = @'
Public Sub Leaf(a() As Long)
    Dim t As Long
    t = a(0)
End Sub

Public Sub Drive()
    Dim a() As Long
    Dim i As Long
    ReDim a(0 To 1999)
    For i = 0 To 1999
        a(i) = i
    Next i
    For i = 1 To 400
        Leaf a
    Next i
End Sub
'@

function Invoke-RingCase {
    # Runs <Leaf>!Drive under the given ring and returns every count the cases compare.
    # The book is built by the test itself: a Skip inside an assigned call would be lost.
    param($Sx, [string]$Leaf, [bool]$RenderArgs, [string]$BufferSize, [string]$WhenFull = 'DROP')
    $log = (Get-XRayPaths $Sx.ProcId).Log

    [void](Set-XRayTraceParam $Sx 'VBA' 'ARGS' $RenderArgs)
    [void](Set-XRayTraceParam $Sx 'VBA' 'RETVAL' $false)
    [void](Set-XRayTraceParam $Sx 'XLL' 'DEPTH' 'OFF')
    $bufferEcho = Set-XRayTraceParam $Sx 'BUFFERSIZE' $BufferSize
    $whenFullEcho = Set-XRayTraceParam $Sx 'BUFFERWHENFULL' $WhenFull

    $mark = Get-LogLength $log
    [void](Invoke-XRayCommand $Sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { throw "did not arm: $armLine" }

    [void]$Sx.App.Run($Leaf + '!Drive')
    $disMark = Get-LogLength $log
    $disarmDrops = Invoke-XRayDisarm $Sx
    [void](Wait-LogLine $log 'VBA tracing: disarmed' $disMark)

    $rows = @(Read-TraceRows $Sx.ProcId)      # the reader throws on any contract breach
    $entryExit = @($rows | Where-Object { $_.kind -eq 'entry' -or $_.kind -eq 'exit' })

    # A dropped row leaves a hole in the producer's input sequence. Interior holes
    # only: a drop after the last written row leaves none, so Disarm's count rules.
    $inputs = @($rows | ForEach-Object { [int64]$_.input } | Sort-Object)
    $holes = if ($inputs.Count) { ($inputs[-1] - $inputs[0] + 1) - $inputs.Count } else { 0 }

    # framesOpened/Closed are tallied in memory and never routed through the ring
    $all = @(Get-Content $log)
    $tail = if ($all.Count -gt $disMark) { $all[$disMark..($all.Count - 1)] } else { @() }
    $totals = $tail | Where-Object { $_ -match 'framesOpened=' } | Select-Object -Last 1
    $ringLine = $tail | Where-Object { $_ -match 'output ring:' } | Select-Object -Last 1
    $opened = if ($totals -match 'framesOpened=(\d+)') { [int]$Matches[1] } else { -1 }
    $closed = if ($totals -match 'framesClosed=(\d+)') { [int]$Matches[1] } else { -1 }
    $lineDrops = if ($ringLine -match '(\d+) dropped') { [int]$Matches[1] } else { -1 }
    $pauses = if ($ringLine -match '(\d+) hot-path pause') { [int]$Matches[1] } else { -1 }

    return [pscustomobject]@{
        BufferEcho = $bufferEcho; WhenFullEcho = $whenFullEcho
        Rows = $rows; EntryExit = $entryExit; Holes = $holes
        DisarmDrops = $disarmDrops; LineDrops = $lineDrops; Pauses = $pauses
        Opened = $opened; Closed = $closed; Expected = $opened + $closed
    }
}
