# End on VBA's error dialog fires no opcode, so its frames are closed later: at disarm, or when
# the next macro starts. Either way the raiser reads `threw` and the frames beneath `abandoned`.
# A chain from a cell keeps its error, which went to the cell.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Function U_Raise(ByVal x As Double) As Double
    Dim t As Long
    t = 1
    Err.Raise 5, "U", "raised"
End Function
Public Function U_Mid(ByVal x As Double) As Double
    U_Mid = U_Raise(x)
End Function
Public Sub M_Raise()
    Dim v As Double
    v = U_Mid(1)
End Sub
Public Sub M_Evaluate()
    Dim v As Variant
    v = Application.Evaluate("U_Mid(1)")
    v = 2
End Sub
Public Sub Noop()
    Dim n As Long
    n = 1
End Sub
'@

# The first activation's exits, in the order they closed.
function FirstChain($Rows, [string[]]$Fns) {
    $out = @()
    foreach ($fn in $Fns) {
        $x = @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $fn })
        $out += $(if ($x.Count) { "$fn=$($x[0].outcome)/$($x[0].trust)" } else { "$fn=(no row)" })
    }
    $out -join ' '
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    New-XRayMacroBook $sx 'DialogEnd' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf
    $ws = $book.Sheet
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')
    $dlgBefore = @(Get-SessionDialogs).Count

    $chain = 'U_Raise', 'U_Mid', 'M_Raise'
    $want = 'U_Raise=threw/{0} U_Mid=abandoned/{0} M_Raise=abandoned/{0}'
    $raise = { try { $app.Run($leaf + '!M_Raise') | Out-Null } catch {} }

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { & $raise }
    $got = FirstChain $s.Rows $chain
    Check 'closed-at-disarm' ($got -eq ($want -f 'flush')) $got

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { & $raise; $app.Run($leaf + '!Noop') | Out-Null }
    $got = FirstChain $s.Rows $chain
    Check 'closed-when-another-macro-starts' ($got -eq ($want -f 'backstop')) $got

    # Run again, the new chain starts at the dead one's stack pointer.
    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { & $raise; & $raise }
    $got = FirstChain $s.Rows $chain
    Check 'closed-when-the-same-macro-runs-again' ($got -eq ($want -f 'backstop')) $got

    # Evaluate gives the function a calling cell, so its error goes there; the macro that
    # called Evaluate is what the dialog ended.
    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { try { $app.Run($leaf + '!M_Evaluate') | Out-Null } catch {}
                                                         $app.Run($leaf + '!Noop') | Out-Null }
    $mid = @($s.Rows | Where-Object { $_.kind -eq 'exit' -and $_.function -eq 'U_Mid' } | ForEach-Object { $_.outcome }) -join ','
    $got = FirstChain $s.Rows 'M_Evaluate'
    Check 'evaluate-errors-go-to-the-cell' ($mid -and ($mid -split ',' | Where-Object { $_ -ne 'unhandled' }).Count -eq 0) "U_Mid=$mid"
    Check 'the-macro-calling-evaluate-reads-abandoned' ($got -eq 'M_Evaluate=abandoned/backstop') $got

    # A cell's chain dies whole too, with no dialog: it keeps its error.
    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { $ws.Range('A1').Formula = '=U_Mid(1)'; $app.Calculate()
                                                         $app.Run($leaf + '!Noop') | Out-Null }
    $got = FirstChain $s.Rows 'U_Raise', 'U_Mid'
    Check 'a-cell-chain-keeps-its-error' ($got -eq 'U_Raise=threw/backstop U_Mid=unhandled/backstop') $got

    if (@(Get-SessionDialogs).Count -gt $dlgBefore) { Write-DialogsHandled }
    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'the error dialog''s End reads threw then abandoned, at disarm or at the next macro'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
