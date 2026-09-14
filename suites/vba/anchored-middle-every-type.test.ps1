# TWO KNOWN PARAMETERS, AND EVERY DECLARED TYPE BETWEEN THEM -- in the gate.
#
# Three parameters. The first and third are always `ByVal As Long` carrying 1001
# and 3003. The MIDDLE one is swept through every type VBA can declare, as both
# a Sub and a Function.
#
# WHAT THIS ADDS OVER `parameter-orders`, which is the neighbouring case and
# already exercises the three-slot ByVal Variant first, middle and last. Three
# things, and the third is the one that matters most:
#
#   1. EVERY middle type, not a chosen permutation set -- Object, a class type,
#      a whole array, a UDT, an Enum, LongPtr, Date, and the ByRef forms.
#   2. Sub AND Function for each, because a Function has a result slot at
#      [R14-8] and a Sub does not (`sub-local-read-as-result` came from exactly
#      that confusion).
#   3. THE ASSERTION CARRIES NO TABLE OF EXPECTED TYPE NAMES. `parameter-orders`
#      states the reported name for each type, which is right for what it tests
#      and is also the part that rots when a type table changes. Here the
#      invariant is a RELATIONSHIP: whatever sits in the middle, the first
#      parameter reads `a1:Long=1001` and the last reads `Long=3003`. That
#      survives every future change to the type table without being edited.
#
# WHY ANCHORS FIND SLOT SLIP AT ALL. The argument label is a FRAME SLOT index
# and slots are not parameters: a `ByVal Variant` is a 24-byte VARIANT occupying
# THREE, so the third parameter lands at `a5`. Read slot 3 instead of slot 5 and
# the answer is still *a number* -- which is why a value that carries no
# position cannot catch it. 1001 and 3003 differ on purpose: two anchors holding
# the same value would let the decoder read one where the other belongs.
#
# THREE PARAMETERS SHOULD DECODE WITHOUT A RESYNCHRONISATION, ALWAYS. These
# bodies are four statements with no call in them. If the walk cannot cross one
# from offset 0 to the exit, a pinned length is wrong somewhere -- so `)~` is a
# failure here, not an acceptable outcome, and neither is a `?` in a body that
# reads all three parameters.
#
# ParamArray is absent: it must be the LAST parameter, so it cannot be a middle.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# Id, the middle's declaration, the driver's local, how it is initialised, and
# how the body READS it. Every body reads all three parameters: an unread one
# emits no typed load and would come back `?` for a reason that is not a defect.
$Middles = @(
  @{ Id='Byte';     Decl='ByVal m As Byte';      Local='As Byte';      Init='= 7'          }
  @{ Id='Integer';  Decl='ByVal m As Integer';   Local='As Integer';   Init='= 7'          }
  @{ Id='Long';     Decl='ByVal m As Long';      Local='As Long';      Init='= 7'          }
  @{ Id='LongLong'; Decl='ByVal m As LongLong';  Local='As LongLong';  Init='= 7'          }
  @{ Id='LongPtr';  Decl='ByVal m As LongPtr';   Local='As LongPtr';   Init='= 7'          }
  @{ Id='Single';   Decl='ByVal m As Single';    Local='As Single';    Init='= 1.5'        }
  @{ Id='Double';   Decl='ByVal m As Double';    Local='As Double';    Init='= 2.5'        }
  @{ Id='Currency'; Decl='ByVal m As Currency';  Local='As Currency';  Init='= 3.5'        }
  @{ Id='Date';     Decl='ByVal m As Date';      Local='As Date';      Init='= #1/1/2020#' }
  @{ Id='Boolean';  Decl='ByVal m As Boolean';   Local='As Boolean';   Init='= True'       }
  @{ Id='String';   Decl='ByVal m As String';    Local='As String';    Init='= "mid"'      }
  @{ Id='Variant';  Decl='ByVal m As Variant';   Local='As Variant';   Init='= 4242'       }
  @{ Id='Enum';     Decl='ByVal m As XREKind';   Local='As XREKind';   Init='= XREOne'     }
  @{ Id='Object';   Decl='ByVal m As Object';    Local='As Object';    Init='Set= New XRCThing'; Obj=$true }
  @{ Id='Class';    Decl='ByVal m As XRCThing';  Local='As XRCThing';  Init='Set= New XRCThing'; Obj=$true }
  @{ Id='RVariant'; Decl='ByRef m As Variant';   Local='As Variant';   Init='= 4242'       }
  @{ Id='RDouble';  Decl='ByRef m As Double';    Local='As Double';    Init='= 2.5'        }
  @{ Id='RString';  Decl='ByRef m As String';    Local='As String';    Init='= "mid"'      }
  # ByRef-only forms: an array or a UDT cannot be ByVal (MS-VBAL 5.3.1.5).
  @{ Id='Array';    Decl='ByRef m() As Long';    Local='(0 To 2) As Long'; Init='(0)= 5'; Arr=$true }
  @{ Id='Udt';      Decl='ByRef m As XRTPoint';  Local='As XRTPoint';  Init='.X= 9'; Udt=$true }
)

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Public Enum XREKind')
    [void]$sb.AppendLine('    XREOne = 1')
    [void]$sb.AppendLine('End Enum')
    [void]$sb.AppendLine('Public Type XRTPoint')
    [void]$sb.AppendLine('    X As Double')
    [void]$sb.AppendLine('    Y As String')
    [void]$sb.AppendLine('End Type')
    [void]$sb.AppendLine('Public gN As Double')
    [void]$sb.AppendLine('Public gV As Variant')

    $cases = @()
    foreach ($m in $Middles) {
        foreach ($kind in @('S','F')) {
            $name = "AM_$($m.Id)_$kind"
            $readMid = if ($m.Obj)     { 'If m Is Nothing Then gN = 1' }
                       elseif ($m.Arr) { 'gN = UBound(m)' }
                       elseif ($m.Udt) { 'gN = m.X' }
                       else            { 'gV = m' }
            if ($kind -eq 'S') { [void]$sb.AppendLine("Public Sub $name(ByVal a As Long, $($m.Decl), ByVal b As Long)") }
            else               { [void]$sb.AppendLine("Public Function $name(ByVal a As Long, $($m.Decl), ByVal b As Long) As Long") }
            [void]$sb.AppendLine('    gN = a')
            [void]$sb.AppendLine("    $readMid")
            [void]$sb.AppendLine('    gN = b')
            if ($kind -eq 'F') { [void]$sb.AppendLine("    $name = a + b") }
            [void]$sb.AppendLine("End $(if ($kind -eq 'S') { 'Sub' } else { 'Function' })")
            $cases += [pscustomobject]@{ Name=$name; Mid=$m.Id; Kind=$kind }
        }
    }

    [void]$sb.AppendLine('Public Sub Drive()')
    [void]$sb.AppendLine('    On Error Resume Next')
    [void]$sb.AppendLine('    Dim r As Long')
    foreach ($m in $Middles) {
        $v = "v_$($m.Id)"
        if ($m.Arr)      { [void]$sb.AppendLine("    Dim $v$($m.Local)"); [void]$sb.AppendLine("    $v$($m.Init -replace '^\(0\)=','(0) =')") }
        elseif ($m.Udt)  { [void]$sb.AppendLine("    Dim $v $($m.Local)"); [void]$sb.AppendLine("    $v$($m.Init -replace '^\.X=','.X =')") }
        elseif ($m.Obj)  { [void]$sb.AppendLine("    Dim $v $($m.Local)"); [void]$sb.AppendLine("    Set $v = New XRCThing") }
        else             { [void]$sb.AppendLine("    Dim $v $($m.Local)"); [void]$sb.AppendLine("    $v $($m.Init)") }
        foreach ($kind in @('S','F')) {
            $n = "AM_$($m.Id)_$kind"
            if ($kind -eq 'S') { [void]$sb.AppendLine("    $n 1001, $v, 3003") }
            else               { [void]$sb.AppendLine("    r = $n(1001, $v, 3003)") }
        }
    }
    [void]$sb.AppendLine('End Sub')
    $moduleCode = $sb.ToString()

    New-XRayMacroBook $sx 'Anchor' @(
        @{ Kind=2; Name='XRCThing'; Code='Public V As Double' }
        @{ Kind=1; Name='AnchorCase'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $entry = Get-FirstEntryByName $rows

    Write-Output ''
    Write-Output 'every declared type, between two known Longs:'
    foreach ($c in $cases) {
        $r = if ($entry.ContainsKey($c.Name)) { $entry[$c.Name] } else { $null }
        if (-not $r) { Write-Output ("  {0,-10} {1}  (no row)" -f $c.Mid, $c.Kind); continue }
        Write-Output ("  {0,-10} {1}  argc={2}  {3,-22} {4}" -f $c.Mid, $c.Kind, $r.argcount, $r.typetext, $r.args)
    }

    $missing = @($cases | Where-Object { -not $entry.ContainsKey($_.Name) } | ForEach-Object { $_.Name })
    Check 'every-middle-type-was-traced' ($missing.Count -eq 0) `
          ("missing: " + $(if ($missing.Count) { $missing -join ',' } else { 'none' }))

    # ---- THE ANCHORS. No expected-type table needed. ----------------------
    $lost = @()
    foreach ($c in $cases) {
        if (-not $entry.ContainsKey($c.Name)) { continue }
        $a = [string]$entry[$c.Name].args
        if ($a -notmatch 'a1:Long=1001')        { $lost += "$($c.Name): first anchor lost -- $a" }
        if ($a -notmatch '(^|\s)a\d+:Long=3003') { $lost += "$($c.Name): LAST anchor lost -- $a" }
    }
    Check 'both-anchors-survive-every-middle-type' ($lost.Count -eq 0) `
          ($(if ($lost.Count) { $lost -join ' | ' } else { "$($cases.Count) cases, both anchors intact" }))

    # ---- argcount is PARAMETERS, not slots: three, always -----------------
    $badCount = @($cases | Where-Object { $entry.ContainsKey($_.Name) -and [int]$entry[$_.Name].argcount -ne 3 } |
                  ForEach-Object { "$($_.Name)=$($entry[$_.Name].argcount)" })
    Check 'argcount-is-three-whatever-the-middle-costs-in-slots' ($badCount.Count -eq 0) `
          ($(if ($badCount.Count) { $badCount -join ' ' } else { 'all 3' }))

    # ---- A THREE-PARAMETER BODY MUST DECODE WITHOUT RESYNCHRONISING -------
    $resynced = @($cases | Where-Object { $entry.ContainsKey($_.Name) -and $entry[$_.Name].typetext -match '~$' } |
                  ForEach-Object { "$($_.Name)=$($entry[$_.Name].typetext)" })
    Check 'no-resynchronisation-on-a-four-statement-body' ($resynced.Count -eq 0) `
          ($(if ($resynced.Count) { $resynced -join ' ' } else { 'every walk reached the exit cleanly' }))

    # A `?` cannot mean "never read" here: the body reads all three.
    $untyped = @($cases | Where-Object { $entry.ContainsKey($_.Name) -and $entry[$_.Name].typetext -match '\?' } |
                 ForEach-Object { "$($_.Name)=$($entry[$_.Name].typetext)" })
    Check 'every-parameter-typed-because-the-body-reads-all-three' ($untyped.Count -eq 0) `
          ($(if ($untyped.Count) { $untyped -join ' ' } else { 'no untyped position' }))

    # ---- WHAT EACH MIDDLE COST IN SLOTS, reported not asserted ------------
    # The slot-consumption table is the thing being learnt, and a ByVal Variant
    # taking three is the case the anchors exist for.
    Write-Output ''
    Write-Output 'slots consumed by the middle (from where 3003 landed):'
    foreach ($c in $cases) {
        if (-not $entry.ContainsKey($c.Name)) { continue }
        if ([string]$entry[$c.Name].args -match 'a(\d+):Long=3003') {
            Write-Output ("  {0,-10} {1}  last anchor at a{2} -> middle occupied {3} slot(s)" -f `
                $c.Mid, $c.Kind, $Matches[1], ([int]$Matches[1] - 2))
        }
    }

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Write-Output ''; Write-Output '---- generated source ----'; Write-Output $moduleCode }


    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "$($cases.Count) cases, both anchors intact, no resynchronisation"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
