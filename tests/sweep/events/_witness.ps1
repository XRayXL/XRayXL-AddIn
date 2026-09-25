# A workbook whose VBA records Excel's events itself, through WithEvents, as the witness the trace is
# compared against. The UDF notes its own calls, so calls and events land in one sequence.
$script:WitnessClass = @'
Public WithEvents App As Application
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Note "SheetChange|" & Target.Address(False, False, xlA1, True) & "|" & Sh.Name
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
    Note "SheetCalculate|" & Sh.Name
End Sub
Private Sub App_AfterCalculate()
    Note "AfterCalculate"
End Sub
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Note "SheetSelectionChange|" & Target.Address(False, False, xlA1, True) & "|" & Sh.Name
End Sub
'@

$script:WitnessModule = @'
Public gLog As String
Public gApp As CApp
Public Sub Note(s As String)
    gLog = gLog & s & ";"
End Sub
Public Function Twice(x As Double) As Double
    Note "Twice"
    Twice = x * 2
End Function
Public Sub Hook()
    Set gApp = New CApp
    Set gApp.App = Application
End Sub
Public Sub ClearLog()
    gLog = ""
End Sub
Public Function GetLog() As String
    GetLog = gLog
End Function
'@

function New-WitnessBook($Sx, [string]$Stem) {
    New-XRayMacroBook $Sx $Stem @(
        @{ Kind = 2; Name = 'CApp'; Code = $script:WitnessClass },
        @{ Kind = 1; Name = 'M';    Code = $script:WitnessModule }
    ) -Cells @{ 'A1' = '3'; 'B1' = '=Twice(A1)' }
    $b = Get-XRayMacroBook $Stem
    [void]$Sx.App.Run("'$($b.Leaf)'!M.Hook")
    return $b
}

# The witness log as entries: @{ Name; Where; Sheet }.
function Get-WitnessLog($Sx, $Book) {
    $text = [string]$Sx.App.Run("'$($Book.Leaf)'!M.GetLog")
    @($text -split ';' | Where-Object { $_ } | ForEach-Object {
        $p = $_ -split '\|'
        [pscustomobject]@{ Name = $p[0]; Where = $(if ($p.Count -gt 1) { $p[1] } else { '' }); Sheet = $(if ($p.Count -gt 2) { $p[2] } else { '' }) }
    })
}

# The trace's rows after the witness log was cleared, which is where the two sequences start together.
function Get-RowsAfterClear($Rows) {
    $clear = @($Rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq 'ClearLog' }) | Select-Object -Last 1
    if (-not $clear) { return @() }
    @($Rows | Where-Object { [int64]$_.seq -gt [int64]$clear.seq })
}
