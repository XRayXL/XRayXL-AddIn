# A RENDERED ARRAY MUST SAY WHERE IT ENDS.
#
# Both columns build an array's text into a FIXED buffer and let it truncate.
# That was harmless while a result cell showed four elements at %g -- sixty
# characters into 256. Two changes made in one sitting removed the headroom
# without either of them looking like it touched arrays:
#
#   * the element cap became 64, shared by both columns, because the SAME array
#     showed 64 elements as an argument and 4 as a result;
#   * doubles moved to %.15g, because 1234567.89012345 was arriving as
#     1.23457e+06.
#
# Sixty-four doubles at fifteen significant digits is roughly 1300 characters.
# The buffer is 256, and it is the CLOSING `}` and the `,...` truncation marker
# that get appended LAST -- so they are the first things lost. What comes out
# reads as a complete array that happens to end:
#
#   Double[1..64]{1.123456789012,2.123456789012,...,17.12345678
#
# A truncated array that ANNOUNCES it is truncated is fine; this project shows
# 64 of a million elements routinely. An array that stops mid-number with no
# marker is a confident wrong answer, and the rule is that a field is written
# empty rather than wrong.
#
# SO THE ASSERTION IS SHAPE, NOT CONTENT. Every rendered array must open with a
# type and bounds, and must CLOSE -- and if it does not carry all of its
# elements it must say so. Both columns are checked on the same arrays, because
# they have separate buffers and each has to hold on its own.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# 64 elements is the cap both columns share, so this is the widest thing they
# are ever asked to print -- the worst case, not an extreme one.
$moduleCode = @'
' Wide DOUBLES: %.15g gives about fourteen characters each.
Public Function WideDblRet() As Double()
    Dim a(1 To 64) As Double
    Dim i As Long
    For i = 1 To 64
        a(i) = i + 0.123456789012
    Next i
    WideDblRet = a
End Function

Public Sub WideDblArg(a() As Double)
    Dim t As Double
    t = a(1)
End Sub

' Wide STRINGS: quoted, so wider still, and a different element decoder.
Public Function WideStrRet() As String()
    Dim a(1 To 64) As String
    Dim i As Long
    For i = 1 To 64
        a(i) = "element-number-" & CStr(i)
    Next i
    WideStrRet = a
End Function

Public Sub WideStrArg(a() As String)
    Dim t As String
    t = a(1)
End Sub

' The CONTROL. Three narrow elements fit any buffer, so if this one is also
' unclosed the fault is in the renderer generally and not in the width.
Public Function NarrowRet() As Long()
    Dim a(1 To 3) As Long
    a(1) = 1: a(2) = 2: a(3) = 3
    NarrowRet = a
End Function

Public Sub NarrowArg(a() As Long)
    Dim t As Long
    t = a(1)
End Sub

Public Sub Drive()
    Dim d(1 To 64) As Double
    Dim s(1 To 64) As String
    Dim n(1 To 3) As Long
    Dim i As Long
    For i = 1 To 64
        d(i) = i + 0.123456789012
        s(i) = "element-number-" & CStr(i)
    Next i
    n(1) = 1: n(2) = 2: n(3) = 3

    WideDblArg d
    WideStrArg s
    NarrowArg n

    Dim r As Variant
    r = WideDblRet()
    r = WideStrRet()
    r = NarrowRet()
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'WideArr' @(
        @{ Kind=1; Name='WideArrCase'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'   'TRUE')
    [void](Set-XRayTraceParam $sx 'VBA' 'RETVAL' 'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH'  'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)

    # Every rendered array, whichever column produced it, as one list.
    $seen = @(
        @{ Where='arg'; Fn='WideDblArg'; Text=(ArgsOf $rows 'WideDblArg'); Wide=$true  }
        @{ Where='arg'; Fn='WideStrArg'; Text=(ArgsOf $rows 'WideStrArg'); Wide=$true  }
        @{ Where='arg'; Fn='NarrowArg';  Text=(ArgsOf $rows 'NarrowArg');  Wide=$false }
        @{ Where='ret'; Fn='WideDblRet'; Text=(RetOf $rows 'WideDblRet'); Wide=$true  }
        @{ Where='ret'; Fn='WideStrRet'; Text=(RetOf $rows 'WideStrRet'); Wide=$true  }
        @{ Where='ret'; Fn='NarrowRet';  Text=(RetOf $rows 'NarrowRet');  Wide=$false }
    )

    Write-Output ''
    Write-Output 'rendered arrays, both columns (tail shown so the ending is visible):'
    foreach ($s in $seen) {
        $t = $s.Text
        $tail = if ($t.Length -gt 70) { '...' + $t.Substring($t.Length - 60) } else { $t }
        Write-Output ("  {0,-3} {1,-12} len={2,-5} {3}" -f $s.Where, $s.Fn, $t.Length, $tail)
    }

    # ---- it reached the decoder at all -------------------------------------
    # An array the walk refused renders as a raw qword, which would make every
    # assertion below vacuously true. This says so first.
    $missing = @($seen | Where-Object { $_.Text -notmatch '\[-?\d+\.\.-?\d+\]\{' })
    Check 'every-array-was-decoded-as-an-array' ($missing.Count -eq 0) `
          ("not rendered as an array: " + $(if ($missing.Count) { (@($missing | ForEach-Object { "$($_.Where)/$($_.Fn)='$($_.Text)'" }) -join ' | ') } else { 'none' }))

    # ---- THE CASE THIS FILE EXISTS FOR: it closes --------------------------
    $unclosed = @($seen | Where-Object { $_.Text -and ($_.Text -notmatch '\}$') })
    Check 'every-rendered-array-closes-its-brace' ($unclosed.Count -eq 0) `
          ("unclosed: " + $(if ($unclosed.Count) { (@($unclosed | ForEach-Object { "$($_.Where)/$($_.Fn) ends '$($_.Text.Substring([Math]::Max(0,$_.Text.Length-24)))'" }) -join ' | ') } else { 'none' }))

    # ---- and an incomplete one says it is incomplete ------------------------
    #
    # Counting elements by commas is only valid once the brace check passes,
    # so this is a separate case: a full 64 needs no marker, anything short
    # of it does.
    $silent = @()
    foreach ($s in $seen) {
        if (-not $s.Text) { continue }
        if ($s.Text -notmatch '\{(.*)\}$') { continue }        # unclosed, reported above
        $body = $Matches[1]
        $want = if ($s.Wide) { 64 } else { 3 }
        $has  = @($body -split ',').Count
        $marked = $body -match '\.\.\.'
        if (($has -lt $want) -and (-not $marked)) {
            $silent += ("{0}/{1} shows {2} of {3} with no marker" -f $s.Where, $s.Fn, $has, $want)
        }
    }
    Check 'a-truncated-array-says-it-was-truncated' ($silent.Count -eq 0) `
          ("silently short: " + $(if ($silent.Count) { $silent -join ' | ' } else { 'none' }))

    # ---- the control -------------------------------------------------------
    # Three Longs fit anything. If these are unclosed the fault is not width.
    # The ARGUMENTS column prefixes its slot -- `a1:Ref&=Long[1..3]{1,2,3}` --
    # so the array text is matched where it sits rather than anchored to the
    # start of the field. (Anchoring it was this file's own first bug, and it
    # failed the control while the product was right.)
    $ctl = @($seen | Where-Object { -not $_.Wide })
    $ctlBad = @($ctl | Where-Object { $_.Text -notmatch 'Long\[1\.\.3\]\{1,2,3\}$' })
    Check 'the-narrow-control-renders-whole' ($ctlBad.Count -eq 0) `
          ("control: " + (@($ctl | ForEach-Object { "$($_.Where)='$($_.Text)'" }) -join ' | '))

    # ---- and the two columns agree about the same array ---------------------
    # Same declaration, same bounds, same element type -- read twice by two
    # decoders. They have disagreed before (bounds, element caps, the BSTR
    # limits), each time silently.
    $pairs = @(
        @{ Name='Double'; A=(ArgsOf $rows 'WideDblArg'); R=(RetOf $rows 'WideDblRet') }
        @{ Name='String'; A=(ArgsOf $rows 'WideStrArg'); R=(RetOf $rows 'WideStrRet') }
        @{ Name='Long';   A=(ArgsOf $rows 'NarrowArg');  R=(RetOf $rows 'NarrowRet')  }
    )
    $disagree = @()
    foreach ($p in $pairs) {
        $ha = if ($p.A -match '([A-Za-z]+\[-?\d+\.\.-?\d+\])') { $Matches[1] } else { '(none)' }
        $hr = if ($p.R -match '([A-Za-z]+\[-?\d+\.\.-?\d+\])') { $Matches[1] } else { '(none)' }
        if ($ha -ne $hr) { $disagree += ("{0}: arg said {1}, ret said {2}" -f $p.Name, $ha, $hr) }
    }
    Check 'both-columns-give-the-same-header-for-the-same-array' ($disagree.Count -eq 0) `
          ("disagreements: " + $(if ($disagree.Count) { $disagree -join ' | ' } else { 'none' }))

    # ---- and they truncate at comparable points -----------------------------
    #
    # The same array should show a similar number of elements in both columns;
    # nothing in a row says the two columns had different room.
    #
    # Not equality: the columns carry different surrounding text, so a few
    # elements either way is structure rather than drift. A four-fold gap is
    # drift.
    $lopsided = @()
    foreach ($p in $pairs) {
        $na = if ($p.A -match '\((\d+) of \d+ shown\)') { [int]$Matches[1] } elseif ($p.A -match '\{(.*)\}') { @($Matches[1] -split ',').Count } else { 0 }
        $nr = if ($p.R -match '\((\d+) of \d+ shown\)') { [int]$Matches[1] } elseif ($p.R -match '\{(.*)\}') { @($Matches[1] -split ',').Count } else { 0 }
        if ($na -lt 1 -or $nr -lt 1) { $lopsided += ("{0}: arg {1}, ret {2}" -f $p.Name, $na, $nr); continue }
        $ratio = [Math]::Max($na, $nr) / [Math]::Min($na, $nr)
        if ($ratio -gt 2) { $lopsided += ("{0}: arg showed {1}, ret showed {2}" -f $p.Name, $na, $nr) }
    }
    Check 'the-two-columns-show-comparable-numbers-of-elements' ($lopsided.Count -eq 0) `
          ("lopsided: " + $(if ($lopsided.Count) { $lopsided -join ' | ' } else { 'none' }))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("{0} rendered arrays, all closed" -f $seen.Count)
}
catch {
    Complete-Test -Fail -Detail ("threw: " + $_.Exception.Message)
}
