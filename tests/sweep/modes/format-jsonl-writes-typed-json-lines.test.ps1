# FORMAT=JSONL writes one typed JSON object a line, a Double included, which the text format leaves
# bare. Every line must parse, `seq` must be dense, and values keep the types VBA and Excel gave them.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    $m = @'
Public Function J_Take(ByVal v As Variant, ByVal d As Double) As Variant
    ' Arithmetic, so the body loads d by its type and the trace can name it.
    Dim e As Double
    e = d * 1
    J_Take = Array(CLng(7), e, "x")
End Function

Public Function J_Containers() As Long
    Dim c As New Collection, d As Object
    c.Add 1.5
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", Nothing
    J_See c
    J_See d
    J_Containers = 1
End Function

Public Sub J_See(ByVal o As Variant)
End Sub
'@
    New-XRayMacroBook $sx 'JsonlFormat' @(@{ Kind = 1; Name = 'M'; Code = $m }) @{
        'A1' = '1.5'; 'B1' = 'x'; 'A2' = '2'; 'B2' = 'TRUE'
        'D1' = '=J_Take(A1:B2,2.5)'
        'H1' = '=TxB(2,3)'
        'J1' = '=J_Containers()'
    }
    $book = Get-XRayMacroBook

    $echo = Set-XRayTraceParam $sx $null 'FORMAT' 'JSONL'
    Check 'format-set-echoes-jsonl' ($echo -match '^FORMAT=JSONL') "echo: $echo"
    Check 'format-reads-back-jsonl' ([string](Get-XRayTraceParam $sx 'FORMAT' $null) -eq 'JSONL') 'GetTraceParam FORMAT'

    [void](Invoke-XRayArmedSession $sx -Body { Invoke-XRayRecalc $app })

    $dir = Join-Path (Get-XRayRoot) 'TraceFiles'
    $file = @(Get-ChildItem (Join-Path $dir ("XRayXL_Trace_*_{0}.jsonl" -f $sx.ProcId)) -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime) | Select-Object -Last 1
    if (-not $file) { Complete-Test -Fail -Detail 'no .jsonl trace was written' }
    Write-Output "trace: $($file.FullName)"

    # ---- every line is one JSON object ---------------------------------------
    $lines = @(Get-Content -LiteralPath $file.FullName -Encoding UTF8 | Where-Object { $_ -ne '' })
    $rows = @(); $bad = @()
    foreach ($l in $lines) {
        try { $rows += ,($l | ConvertFrom-Json) } catch { $bad += $l.Substring(0, [Math]::Min(80, $l.Length)) }
    }
    Check 'every-line-parses-as-json' (($lines.Count -gt 0) -and ($bad.Count -eq 0)) `
          ("{0} line(s), {1} unparsable; first: {2}" -f $lines.Count, $bad.Count, $(if ($bad.Count) { $bad[0] } else { 'none' }))
    Check 'no-csv-header' ($lines.Count -gt 0 -and $lines[0].StartsWith('{')) "first line: $($lines[0].Substring(0, [Math]::Min(60, $lines[0].Length)))"

    $seqOk = $true
    for ($i = 0; $i -lt $rows.Count; $i++) { if ([long]$rows[$i].seq -ne $i + 1) { $seqOk = $false; break } }
    Check 'seq-is-dense-in-file-order' $seqOk 'seq must run 1..N down the file'

    $mine = @($rows | Where-Object { $_.module -like "*$($book.Leaf)*" -or $_.source -eq 'XLL' })
    $e = @($mine | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'J_Take' }) | Select-Object -First 1
    $x = @($mine | Where-Object { $_.kind -eq 'exit'  -and $_.function -eq 'J_Take' }) | Select-Object -First 1
    if (-not $e -or -not $x) { Complete-Test -Fail -Detail 'J_Take was not traced on both sides' }

    # ---- the Range argument: an object whose value is a 2-D grid of typed cells ----
    $a1 = $e.args | Where-Object { $_.slot -eq 1 }
    $rng = $a1.value
    Check 'range-arg-is-an-object' (($a1.type -eq 'Variant') -and ($rng.t -eq 'Object') -and ($rng.class -eq 'Range')) `
          ("slot 1: type={0} t={1} class={2}" -f $a1.type, $rng.t, $rng.class)
    $grid = $rng.value
    $boundsOk = ($grid.t -eq 'Array') -and (@($grid.bounds).Count -eq 2) -and
                ((@($grid.bounds[0]) -join '..') -eq '1..2') -and ((@($grid.bounds[1]) -join '..') -eq '1..2')
    Check 'range-value-states-both-bounds' $boundsOk ("t={0} bounds={1}" -f $grid.t, ($grid.bounds | ConvertTo-Json -Compress))
    $r1 = @($grid.v[0]); $r2 = @($grid.v[1])
    Check 'row-1-is-first-and-whole' (($r1.Count -eq 2) -and ($r1[0].t -eq 'Double') -and ([double]$r1[0].v -eq 1.5) -and
                                       ($r1[1].t -eq 'String') -and ($r1[1].v -eq 'x')) ($grid.v[0] | ConvertTo-Json -Compress)
    Check 'row-2-types-kept' (($r2.Count -eq 2) -and ($r2[0].t -eq 'Double') -and ([double]$r2[0].v -eq 2) -and
                               ($r2[1].t -eq 'Boolean') -and ($r2[1].v -eq $true)) ($grid.v[1] | ConvertTo-Json -Compress)

    # ---- a Collection is a 1-D array of Variants, a Dictionary a 2-D one of {key,item} ----
    $seen = @($mine | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'J_See' } | ForEach-Object { $_.args[0].value })
    $coll = $seen | Where-Object { $_.class -eq 'Collection' } | Select-Object -First 1
    $dict = $seen | Where-Object { $_.class -eq 'Dictionary' } | Select-Object -First 1
    Check 'collection-value-is-a-1d-array' ($coll -and ($coll.value.t -eq 'Array') -and ($coll.value.elem -eq 'Variant') -and
                                            ((@($coll.value.bounds[0]) -join '..') -eq '1..1') -and
                                            ($coll.value.v[0].t -eq 'Double') -and ([double]$coll.value.v[0].v -eq 1.5)) `
          ($coll | ConvertTo-Json -Compress -Depth 8)
    $pair = if ($dict) { @($dict.value.v[0]) } else { @() }
    Check 'dictionary-value-is-key-item-rows' ($dict -and ($dict.value.t -eq 'Array') -and (@($dict.value.bounds).Count -eq 2) -and
                                               ((@($dict.value.bounds[0]) -join '..') -eq '0..0') -and
                                               ((@($dict.value.bounds[1]) -join '..') -eq '0..1') -and ($pair.Count -eq 2) -and
                                               ($pair[0].t -eq 'String') -and ($pair[0].v -eq 'k') -and ($pair[1].t -eq 'Nothing')) `
          ($dict | ConvertTo-Json -Compress -Depth 8)

    # ---- a declared Double keeps its type descriptor ---------------------------
    $d = $e.args | Where-Object { $_.type -eq 'Double' }
    Check 'double-arg-is-typed' (($d.value.t -eq 'Double') -and ([double]$d.value.v -eq 2.5)) ($d | ConvertTo-Json -Compress -Depth 5)

    # ---- the Variant array result: each element names its type -----------------
    $ret = $x.ret
    $v = @($ret.v)
    Check 'ret-is-a-typed-variant-array' (($x.rettype -eq 'Variant') -and ($ret.t -eq 'Array') -and ($ret.elem -eq 'Variant') -and
                                          ((@($ret.bounds[0]) -join '..') -eq '0..2') -and ($v.Count -eq 3)) ($ret | ConvertTo-Json -Compress -Depth 5)
    if ($v.Count -eq 3) {
        Check 'ret-elements-typed' (($v[0].t -eq 'Long') -and ([long]$v[0].v -eq 7) -and ($v[1].t -eq 'Double') -and
                                    ([double]$v[1].v -eq 2.5) -and ($v[2].t -eq 'String') -and ($v[2].v -eq 'x')) ($ret.v | ConvertTo-Json -Compress)
    }

    # ---- the XLL side writes the same shape ------------------------------------
    $xe = @($mine | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'TxB' }) | Select-Object -First 1
    $xx = @($mine | Where-Object { $_.kind -eq 'exit'  -and $_.function -eq 'TxB' }) | Select-Object -First 1
    Check 'xll-args-are-typed' ($xe -and (@($xe.args).Count -eq 2) -and ($xe.args[0].type -eq 'B') -and
                                ($xe.args[0].value.t -eq 'Double') -and ([double]$xe.args[0].value.v -eq 2)) ($xe.args | ConvertTo-Json -Compress -Depth 5)
    Check 'xll-ret-is-typed' ($xx -and ($null -ne $xx.ret.t)) ($xx.ret | ConvertTo-Json -Compress -Depth 5)
    Check 'integers-are-numbers' ($e.span -is [long] -or $e.span -is [int] -or $e.span -is [decimal]) ("span is " + $e.span.GetType().Name)

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails check(s) failed" }
    Complete-Test -Pass -Detail ("{0} JSON lines, all parsed, values typed" -f $lines.Count)
}
catch {
    Complete-Test -Fail -Detail ("threw: " + $_.Exception.Message)
}
