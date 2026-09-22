# BREAKPOINTS is a VBA setting that adds one column, `breaks`, to the trace file. Off -- the
# default -- the file is exactly as it always was; on, the header gains `breaks` and every VBA
# exit row carries its count, 0 when nothing stopped, while every other row leaves it empty.
# A stop itself needs a person at the editor, so it is not driven here.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # ---- the setting: VBA only, off by default ----------------------------
    $e = Set-XRayTraceParam $sx 'XLL' 'BREAKPOINTS' 'TRUE'
    Check 'xll-breakpoints-refused' ($e -match '^#Err - BREAKPOINTS Parameter only available for VBA') $e
    $e = Set-XRayTraceParam $sx $null 'BREAKPOINTS' 'TRUE'
    Check 'omitted-source-breakpoints-refused' ($e -match 'only available for VBA') $e
    $v = [string](Get-XRayTraceParam $sx 'VBA' 'BREAKPOINTS')
    Check 'vba-breakpoints-default-off' ($v -eq 'FALSE') "VBA BREAKPOINTS='$v'"
    $v = [string](Get-XRayTraceParam $sx 'XLL' 'BREAKPOINTS')
    Check 'get-xll-breakpoints-says-vba-only' ($v -match 'VBA only') "XLL BREAKPOINTS='$v'"

    $srcM = @'
Public Function B_Add(ByVal a As Double, ByVal b As Double) As Double
    B_Add = a + b
End Function
'@
    New-XRayMacroBook $sx 'Breaks' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{
        'A1' = '=B_Add(1,2)'
        'B1' = '=TxB(2,3)'
    }

    function Trace-Once {
        $mark = Get-LogLength $paths.Log
        [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
        [void](Wait-LogLine $paths.Log 'VBA tracing: ' $mark)
        Invoke-XRayRecalc $app
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }
        $csv = Get-XRayTraceCsv $sx.ProcId
        return @{ Header = (Get-Content $csv -TotalCount 1); Rows = @(Read-TraceFile $csv) }
    }

    # ---- off: the file is unchanged ---------------------------------------
    $t = Trace-Once
    Check 'off-header-has-no-breaks' ($t.Header -notmatch 'breaks') $t.Header
    Check 'off-traced-something' ($t.Rows.Count -gt 0) "$($t.Rows.Count) rows"
    Check 'off-rows-have-no-breaks-field' (-not ($t.Rows | Where-Object { $_.PSObject.Properties['breaks'] })) 'a row has a breaks field'

    # ---- on: one more column, a count on VBA exits, empty elsewhere --------
    Clear-XRayStaleTraces $sx.ProcId
    $e = Set-XRayTraceParam $sx 'VBA' 'BREAKPOINTS' 'TRUE'
    Check 'vba-breakpoints-accepted' ($e -notmatch '#Err') $e
    $t = Trace-Once
    Check 'on-header-ends-with-breaks' ($t.Header -match ',trust,breaks$') $t.Header
    $vbaExits = @($t.Rows | Where-Object { $_.source -eq 'VBA' -and $_.kind -eq 'exit' })
    $others   = @($t.Rows | Where-Object { -not ($_.source -eq 'VBA' -and $_.kind -eq 'exit') })
    Check 'on-has-a-vba-exit' ($vbaExits.Count -gt 0) "$($vbaExits.Count) VBA exit rows"
    Check 'on-vba-exits-count-zero' (-not ($vbaExits | Where-Object { $_.breaks -ne '0' })) `
          (($vbaExits | ForEach-Object { "$($_.function)=$($_.breaks)" }) -join ',')
    Check 'on-xll-rows-present' (@($others | Where-Object { $_.source -eq 'XLL' }).Count -gt 0) 'no XLL row to check'
    Check 'on-other-rows-empty' (-not ($others | Where-Object { $_.breaks -ne '' })) `
          (($others | Where-Object { $_.breaks -ne '' } | ForEach-Object { "$($_.source) $($_.kind)=$($_.breaks)" }) -join ',')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'BREAKPOINTS is VBA only and off by default; on, it adds breaks with 0 on VBA exits'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
