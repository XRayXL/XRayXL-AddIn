# The same types in different orders: fixed permutations, in the gate.
#
# The argument label is the frame slot, and a `ByVal Variant` occupies three, so in `(Variant,
# Long)` the Long is `a4` and in `(Long, Variant)` it is `a1`. Wrong slot arithmetic still
# produces something that looks like a value.
#
# Fixed rather than random, so a slot-arithmetic regression fails every run and names the same
# case each time. Each value carries its position (a Long at position 3 holds 1003), so a
# transposition fails.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# The permutations. Each entry is the ordered parameter list of one generated
# Sub: a declared type and a mechanism. The set is chosen so the awkward
# members -- the three-slot ByVal Variant, and the types whose REPORTED name
# differs from the declared one -- appear first, last and in the middle.
$Orders = @(
    @{ N='P1'; P=@('ByVal:Variant','ByVal:Long','ByVal:String') }            # 3-slot first
    @{ N='P2'; P=@('ByVal:Long','ByVal:Variant','ByVal:String') }            # 3-slot middle
    @{ N='P3'; P=@('ByVal:Long','ByVal:String','ByVal:Variant') }            # 3-slot last
    @{ N='P4'; P=@('ByVal:Variant','ByVal:Variant','ByVal:Long') }           # two of them
    @{ N='P5'; P=@('ByRef:Variant','ByVal:Variant','ByRef:Long') }           # 1-slot and 3-slot mixed
    @{ N='P6'; P=@('ByVal:Boolean','ByVal:Variant','ByRef:Boolean') }        # alias around a 3-slot
    @{ N='P7'; P=@('ByRef:LongLong','ByVal:Variant','ByVal:LongLong') }      # Ref& around a 3-slot
    @{ N='P8'; P=@('ByVal:String','ByRef:String','ByVal:Variant','ByRef:Double') }
    @{ N='P9'; P=@('ByVal:Byte','ByVal:Integer','ByVal:Long','ByVal:LongLong','ByVal:Currency') }
    @{ N='PA'; P=@('ByRef:Currency','ByRef:Byte','ByRef:Integer','ByRef:Long','ByRef:Double') }
    @{ N='PB'; P=@('ByVal:Currency','ByRef:Currency','ByVal:Currency') }     # same type, three ways
    @{ N='PC'; P=@('ByRef:String','ByVal:Boolean','ByRef:Variant','ByVal:Single') }
)

# Declared type -> VBA declaration, the value planted at position k, the text
# the decoder must produce, the name it REPORTS, and how many slots it costs.
#
# The reported name is not always the declared one, and that is asserted rather
# than worked around: a Boolean is held in an I2 so it reports as Integer, and a
# ByRef LongLong goes through opcode 747 -- "eight bytes by reference" -- which
# fires for LongLong, LongPtr and arrays alike and so asserts no type.
$TypeTable = @{
    'Long'     = @{ Decl='Long';     Val={param($k) "$(1000+$k)"};         Text={param($k) "$(1000+$k)"};          Rep={param($r) if($r){'Long&'}else{'Long'}} }
    'Integer'  = @{ Decl='Integer';  Val={param($k) "$(100+$k)"};          Text={param($k) "$(100+$k)"};           Rep={param($r) if($r){'Integer&'}else{'Integer'}} }
    'Byte'     = @{ Decl='Byte';     Val={param($k) "$(10+$k)"};           Text={param($k) "$(10+$k)"};            Rep={param($r) if($r){'Byte&'}else{'Byte'}} }
    'Single'   = @{ Decl='Single';   Val={param($k) "$(200+$k).5"};        Text={param($k) "$(200+$k).5"};         Rep={param($r) if($r){'Single&'}else{'Single'}} }
    'Double'   = @{ Decl='Double';   Val={param($k) "$(300+$k).25"};       Text={param($k) "$(300+$k).25"};        Rep={param($r) if($r){'Double&'}else{'Double'}} }
    'Currency' = @{ Decl='Currency'; Val={param($k) "$(400+$k).5"};        Text={param($k) "$(400+$k).5000"};      Rep={param($r) if($r){'Currency&'}else{'Currency'}} }
    'String'   = @{ Decl='String';   Val={param($k) "`"s$k`""};            Text={param($k) "`"s$k`""};             Rep={param($r) if($r){'String&'}else{'String'}} }
    'LongLong' = @{ Decl='LongLong'; Val={param($k) "$(4294967296 + $k)"}; Text={param($k) "$(4294967296 + $k)"};  Rep={param($r) if($r){'Ref&'}else{'LongLong'}} }
    'Boolean'  = @{ Decl='Boolean';  Val={param($k) 'True'};               Text={param($k) '-1'};                  Rep={param($r) if($r){'Integer&'}else{'Integer'}} }
    'Variant'  = @{ Decl='Variant';  Val={param($k) "$(100000+$k)"};       Text={param($k) "Long($(100000+$k))"};  Rep={param($r) if($r){'Variant&'}else{'Variant'}} }
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # ---- build the module from the permutation table ----
    $sb = New-Object System.Text.StringBuilder
    $plan = @()
    foreach ($o in $Orders) {
        $params = @()
        $slot = 1                       # slot 0 is reserved; arguments start at 1
        $k = 0
        foreach ($spec in $o.P) {
            $k++
            $mech, $tn = $spec -split ':'
            $byRef = ($mech -eq 'ByRef')
            $ti = $TypeTable[$tn]
            # A ByVal Variant is a 24-byte VARIANT: THREE slots. A ByRef Variant
            # is one, because the slot holds a pointer.
            $slots = if ((-not $byRef) -and $tn -eq 'Variant') { 3 } else { 1 }
            $params += [pscustomobject]@{
                Idx=$k; TypeName=$tn; T=$ti; ByRef=$byRef; Slot=$slot
                Val=(& $ti.Val $k); Text=(& $ti.Text $k); Label=(& $ti.Rep $byRef)
            }
            $slot += $slots
        }
        $decl = ($params | ForEach-Object {
            $m = if ($_.ByRef) { 'ByRef' } else { 'ByVal' }
            "$m p$($_.Idx) As $($_.T.Decl)" }) -join ', '

        # Every parameter is READ: a parameter the body never reads emits no
        # typed load and has no recoverable type, which would fail this for a
        # reason that is not a defect.
        [void]$sb.AppendLine("Public Sub $($o.N)($decl)")
        foreach ($p in $params) {
            [void]$sb.AppendLine("    Dim t$($p.Idx) As $($p.T.Decl)")
            [void]$sb.AppendLine("    t$($p.Idx) = p$($p.Idx)")
        }
        [void]$sb.AppendLine("End Sub")
        [void]$sb.AppendLine()
        $plan += [pscustomobject]@{ Name=$o.N; Params=$params }
    }

    [void]$sb.AppendLine("Public Sub Drive()")
    foreach ($pr in $plan) {
        foreach ($p in $pr.Params) {
            [void]$sb.AppendLine("    Dim $($pr.Name)_$($p.Idx) As $($p.T.Decl)")
            [void]$sb.AppendLine("    $($pr.Name)_$($p.Idx) = $($p.Val)")
        }
        $args = ($pr.Params | ForEach-Object { "$($pr.Name)_$($_.Idx)" }) -join ', '
        [void]$sb.AppendLine("    $($pr.Name) $args")
    }
    [void]$sb.AppendLine("End Sub")
    $moduleCode = $sb.ToString()

    # ---- OPTION BASE, in its own module ------------------------------------
    #
    # `Dim a(3)` is not one thing. Under the default base 0 it is 0..3 -- FOUR
    # elements -- and under `Option Base 1` it is 1..3, three. A bare count
    # cannot be matched back to the `a(3)` somebody wrote, and that is why both
    # columns state both bounds rather than a count.
    #
    # Two modules, the SAME declaration, so the difference is the directive and
    # nothing else.
    $baseCode = @'
Option Base 1
Public Sub OB1(a() As Long)
    Dim t As Long
    t = a(1)
End Sub
Public Sub OB1Drive()
    Dim a(3) As Long
    a(1) = 11: a(2) = 22: a(3) = 33
    OB1 a
End Sub
'@
    $base0Code = @'
Public Sub OB0(a() As Long)
    Dim t As Long
    t = a(0)
End Sub
Public Sub OB0Drive()
    Dim a(3) As Long
    a(0) = 10: a(1) = 11: a(2) = 12: a(3) = 13
    OB0 a
End Sub
'@

    # one module per Option Base: the directive governs the whole module
    New-XRayMacroBook $sx 'Order' @(
        @{ Kind=1; Name='BaseOne';   Code=$baseCode }
        @{ Kind=1; Name='BaseZero';  Code=$base0Code }
        @{ Kind=1; Name='OrderCase'; Code=$moduleCode }
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
    $app.Run($leaf + '!OB1Drive') | Out-Null
    $app.Run($leaf + '!OB0Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $entry = Get-FirstEntryByName $rows

    Write-Output ''
    Write-Output 'the same types, ordered differently:'
    foreach ($pr in $plan) {
        $shape = ($pr.Params | ForEach-Object { "$(if($_.ByRef){'R'}else{'V'}):$($_.TypeName)" }) -join ' '
        $got = if ($entry.ContainsKey($pr.Name)) { [string]$entry[$pr.Name].args } else { '(no row)' }
        Write-Output ("  {0}  {1,-52} {2}" -f $pr.Name, $shape, $got)
    }

    $missing = @($plan | Where-Object { -not $entry.ContainsKey($_.Name) } | ForEach-Object { $_.Name })
    Check 'every-ordering-was-traced' ($missing.Count -eq 0) `
          ("missing: " + $(if ($missing.Count) { $missing -join ',' } else { 'none' }))

    # ---- argcount counts PARAMETERS, whatever the slots did ----------------
    $badCount = @()
    foreach ($pr in $plan) {
        if (-not $entry.ContainsKey($pr.Name)) { continue }
        $want = $pr.Params.Count
        $got  = [int]$entry[$pr.Name].argcount
        if ($got -ne $want) { $badCount += ("{0}: argcount={1} want {2}" -f $pr.Name, $got, $want) }
    }
    Check 'argcount-is-parameters-not-slots-in-every-order' ($badCount.Count -eq 0) `
          ($(if ($badCount.Count) { $badCount -join ' | ' } else { 'all correct' }))

    # ---- EVERY PARAMETER, AT ITS SLOT, IN EVERY ORDER ----------------------
    $wrong = @(); $checked = 0
    foreach ($pr in $plan) {
        if (-not $entry.ContainsKey($pr.Name)) { continue }
        $argsText = [string]$entry[$pr.Name].args
        foreach ($p in $pr.Params) {
            $checked++
            $want = "a$($p.Slot):$($p.Label)=$($p.Text)"
            if ($argsText -notlike "*$want*") {
                $wrong += ("{0} p{1} ({2}{3}) wanted [{4}] got [{5}]" -f `
                    $pr.Name, $p.Idx, $(if ($p.ByRef){'ByRef '}else{'ByVal '}), $p.TypeName, $want, $argsText)
            }
        }
    }
    Check 'every-parameter-lands-on-its-own-slot-in-every-order' ($wrong.Count -eq 0) `
          ("checked=$checked " + $(if ($wrong.Count) { ($wrong | Select-Object -First 3) -join ' | ' } else { 'all correct' }))

    # The order-specific assertion. P1/P2/P3 are the same three types in three orders; a ByVal
    # Variant costs three slots, so the Long sits at a4, a1 and a1. Naming them means the
    # failure says which order broke.
    $p1 = if ($entry.ContainsKey('P1')) { [string]$entry['P1'].args } else { '' }
    $p2 = if ($entry.ContainsKey('P2')) { [string]$entry['P2'].args } else { '' }
    $p3 = if ($entry.ContainsKey('P3')) { [string]$entry['P3'].args } else { '' }
    Check 'a-three-slot-variant-shifts-what-follows-it' `
          (($p1 -match 'a4:Long=1002') -and ($p2 -match 'a1:Long=1001') -and ($p3 -match 'a1:Long=1001')) `
          ("Variant first: '$p1' | middle: '$p2' | last: '$p3'")

    # ---- OPTION BASE: the same declaration, two meanings --------------------
    Write-Output ''
    Write-Output 'Option Base -- Dim a(3) declared identically in two modules:'
    Write-Output ("  base 0  " + (ArgsOf $rows 'OB0'))
    Write-Output ("  base 1  " + (ArgsOf $rows 'OB1'))

    # THE POINT: a bare count could not tell these apart. `Dim a(3)` is four
    # elements under base 0 and three under base 1, so `[4]` and `[3]` would
    # both be true and neither would match the source.
    Check 'option-base-0-reports-its-real-lower-bound' `
          ((ArgsOf $rows 'OB0') -match '\[0\.\.3\]') `
          ("Dim a(3) under base 0 is 0..3, four elements: " + (ArgsOf $rows 'OB0'))
    Check 'option-base-1-reports-its-real-lower-bound' `
          ((ArgsOf $rows 'OB1') -match '\[1\.\.3\]') `
          ("Dim a(3) under base 1 is 1..3, three elements: " + (ArgsOf $rows 'OB1'))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Write-Output ''; Write-Output '---- generated source ----'; Write-Output $moduleCode }


    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "$($plan.Count) orderings, $checked parameters, every one on its own slot"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
