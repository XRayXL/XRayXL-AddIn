# Proof that binding by window handle lands on exactly the designated Excel, not
# "some Excel": the process-per-test design rests on it.
. (Join-Path $PSScriptRoot '..\..\TestKit.ps1')

try {
    $sx = Connect-TestExcel
    $allGood = $true

    # Connect-TestExcel already refuses a mismatched hwnd; this checks what it returned.
    if (-not [string]::IsNullOrEmpty($env:STRETCH_SESSION_PID)) {
        $ok = ($sx.ProcId -eq [int]$env:STRETCH_SESSION_PID)
        if (-not $ok) { $allGood = $false }
        Write-TestCase -Name 'designated-pid' -Pass:$ok -Fail:(-not $ok) `
                       -Detail "session=$($sx.ProcId) designated=$($env:STRETCH_SESSION_PID)"
    }
    else {
        Write-TestCase -Name 'designated-pid' -Pass -Detail 'standalone -- no designated pid to compare'
    }

    # the object model and the window must agree about which Excel this is
    $appHwnd = [int64]$sx.App.Hwnd
    $ok2 = ($appHwnd -eq [int64]$sx.Hwnd)
    if (-not $ok2) { $allGood = $false }
    Write-TestCase -Name 'app-hwnd-matches' -Pass:$ok2 -Fail:(-not $ok2) `
                   -Detail "app.Hwnd=$appHwnd session.Hwnd=$([int64]$sx.Hwnd)"

    # A round trip proves the proxy is live. Each link is held and released: a dotted chain
    # leaks an RCW per dot, which keeps Excel alive until the close deadline.
    $booksRef = $sx.App.Workbooks
    $bookRef  = $booksRef.Item(1)
    $sheetsRef = $bookRef.Worksheets
    $sheetRef  = $sheetsRef.Item(1)
    $cellRef   = $sheetRef.Range('A1')
    $cellRef.Value2 = 42137
    $readBack = [int]$cellRef.Value2
    $cellRef.Value2 = $null                       # leave the baseline as found
    # a dirty workbook turns the session close into a hidden "Save changes?" dialog
    $bookRef.Saved = $true
    foreach ($r in @($cellRef, $sheetRef, $sheetsRef, $bookRef, $booksRef)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($r)
    }
    $ok3 = ($readBack -eq 42137)
    if (-not $ok3) { $allGood = $false }
    Write-TestCase -Name 'om-roundtrip' -Pass:$ok3 -Fail:(-not $ok3) -Detail "wrote=42137 read=$readBack"

    if ($allGood) { Complete-Test -Pass -Detail "managed=$($sx.Managed)" }
    else          { Complete-Test -Fail -Detail 'one or more identity cases failed' }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
