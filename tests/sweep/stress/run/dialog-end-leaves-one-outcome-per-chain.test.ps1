$case = @{ Name='dialog-end-leaves-one-outcome-per-chain'
     Modules=@{
       'M'=@'
Public Sub RunCross()
    Dim v As Double
    On Error GoTo Caught
    v = Application.Run("'DlgLib.xlam'!LibRaise", -1)
    Exit Sub
Caught:
    v = -1
End Sub
'@
     }
     Deps=@(
       @{ Name='DlgLib'; IsAddin=$true; Modules=@{
         'L'=@'
Public Function LibRaise(ByVal x As Double) As Double
    If x <= 0 Then Err.Raise 1001, "Lib", "x must be positive"
    LibRaise = Sqr(x)
End Function
'@
       } }
     )
     Trigger=@{ Kind='Run'; Name='RunCross'; MayRaise=$true; ExpectDialog=$true }
     Expect={ param($t)
        # An error raised in another project under Application.Run cannot be trapped by the caller,
        # so the watchdog presses End and both frames stay open until disarm closes them. The inner
        # frame's trailer can survive while the outer one's does not; the raiser must still read
        # threw and its caller abandoned, not the other way round.
        if ($t.dialogs -lt 1) { return "expected the modal VBA dialog, none appeared" }
        $x = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and @('RunCross', 'LibRaise') -contains $_.function })
        if ($x.Count -ne 2) { return "expected an exit row for RunCross and LibRaise, got $($x.Count)" }
        $want = @{ LibRaise = 'threw'; RunCross = 'abandoned' }
        $wrong = @($x | Where-Object { $_.outcome -ne $want[$_.function] -or $_.trust -ne 'flush' })
        if ($wrong.Count) {
            return ("expected LibRaise threw and RunCross abandoned, both flush: " +
                    (($x | ForEach-Object { "$($_.function)=$($_.outcome)/$($_.trust)" }) -join ' ')) }
        $null }
     Why='the frames an error dialog''s End leaves open are closed at disarm with the raiser threw and its caller abandoned, never swapped' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
