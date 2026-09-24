# A healthy ring loses nothing and keeps file order: the control the starvation cases are measured
# against. The lean driver makes 800 leaf activations plus its own frame: 1602 entry/exit rows.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_driver.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    New-XRayMacroBook $sx 'RingOk' @(@{ Kind = 1; Name = 'RingCase'; Code = $RingLeanModule })

    # lean rows: this case is about the ring, not the decoder
    $r = Invoke-RingCase $sx -Leaf (Get-XRayMacroBook).Leaf -RenderArgs $false -BufferSize '64'
    Write-Output ("rows=$($r.Rows.Count) entry/exit=$($r.EntryExit.Count) drops=$($r.DisarmDrops) inputHoles=$($r.Holes) framesOpened=$($r.Opened) framesClosed=$($r.Closed)")

    Check 'buffer-set-echoes-ring' ($r.BufferEcho -match 'ring') "echo: $($r.BufferEcho)"
    Check 'the-ring-delivered-every-row' ($r.EntryExit.Count -eq 1602) "$($r.EntryExit.Count) entry/exit rows, expected 1602"
    Check 'no-drops-when-the-buffer-is-ample' ($r.DisarmDrops -eq 0) `
          "XRayXL_Disarm reported $($r.DisarmDrops) dropped row(s) from a 64 MB ring"
    Check 'no-holes-in-the-producer-sequence' ($r.Holes -eq 0) `
          "$($r.Holes) hole(s) in the input column -- rows the producer emitted are not in the file"
    Check 'every-frame-is-in-the-file' (($r.Opened -ge 0) -and ($r.EntryExit.Count -eq $r.Expected)) `
          "file entry/exit=$($r.EntryExit.Count) vs framesOpened+framesClosed=$($r.Expected)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails check(s) failed" }
    Complete-Test -Pass -Detail ("ring 64MB: $($r.EntryExit.Count) rows, 0 dropped")
}
catch {
    Complete-Test -Fail -Detail ("threw: " + $_.Exception.Message)
}
