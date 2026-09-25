$case = @{ Name='dialog-unhandled-error-run'
     Modules=@{
       'M'=@'
Public Sub RaiseLeaf()
    Err.Raise 5901, "RaiseLeaf", "unhandled on purpose"
End Sub
Public Sub RaiseMid()
    RaiseLeaf
End Sub
Public Sub RaiseTop()
    RaiseMid
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='RaiseTop'; MayRaise=$true; ExpectDialog=$true }
     Expect={ param($t)
        # An unhandled Err.Raise under Application.Run raises Excel's modal VBA dialog, which
        # blocks the calling thread until the watchdog presses End. That fires no exit opcode,
        # so disarm finds the three frames dead: the leaf raised, and the two beneath were ended.
        if ($t.dialogs -lt 1) { return "expected the modal VBA dialog, none appeared" }
        if ($t.framesOpened -lt 3) { return "expected >=3 frames, got $($t.framesOpened)" }
        $why = Assert-VbaTraced $t 'RaiseTop','RaiseMid','RaiseLeaf'; if ($why) { return $why }
        $x = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' })
        $want = @{ RaiseLeaf = 'threw'; RaiseMid = 'abandoned'; RaiseTop = 'abandoned' }
        $wrong = @($x | Where-Object { $_.outcome -ne $want[$_.function] -or $_.trust -ne 'flush' })
        if ($x.Count -ne 3 -or $wrong.Count) {
            return ("expected RaiseLeaf threw, RaiseMid and RaiseTop abandoned, all flush; got " +
                    (($x | ForEach-Object { "$($_.function)=$($_.outcome)/$($_.trust)" }) -join ' ')) }
        $null }
     Calls=@(
        @{ Function='RaiseTop';  Depth='1'; Parent=-1 }
        @{ Function='RaiseMid';  Depth='2'; Parent=0 }
        @{ Function='RaiseLeaf'; Depth='3'; Parent=1 } )
     Why='an unhandled error under Application.Run raises Excels modal VBA dialog; the watchdog presses End, which abandons frames without any exit opcode' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
