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
        # so the watchdog presses End and both frames stay open until disarm closes them. A frame
        # beneath a still-running one must be judged running too.
        if ($t.dialogs -lt 1) { return "expected the modal VBA dialog, none appeared" }
        $x = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and @('RunCross', 'LibRaise') -contains $_.function })
        if ($x.Count -ne 2) { return "expected an exit row for RunCross and LibRaise, got $($x.Count)" }
        $flushed = @($x | Where-Object { $_.trust -eq 'flush' })
        $split = @($flushed | Where-Object { $_.outcome -ne 'returned' })
        if ($split.Count) {
            return ("frames closed at disarm read different outcomes: " +
                    (($x | ForEach-Object { "$($_.function)=$($_.outcome)/$($_.trust)" }) -join ' ')) }
        $null }
     Why='the frames an error dialog''s End leaves open are closed at disarm with one outcome, not split into a thrower and a caller that swapped places' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
