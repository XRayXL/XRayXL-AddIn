# PROOF: the by-handle binding lands on EXACTLY the designated Excel.
#
# This is the one mechanism the whole process-per-test design rests on:
# AccessibleObjectFromWindow(OBJID_NATIVEOM) against the hwnd the manager
# handed us must yield the object model of that instance -- not "some Excel".
# Three cases, each a distinct identity check; any one failing fails the test.
. (Join-Path $PSScriptRoot '..\..\TestKit.ps1')

try {
    $sx = Connect-TestExcel
    $allGood = $true

    # 1. The session's pid is the designated pid. (Connect-TestExcel refuses
    #    a mismatched hwnd outright; this asserts what it returned.)
    if (-not [string]::IsNullOrEmpty($env:STRETCH_SESSION_PID)) {
        $ok = ($sx.ProcId -eq [int]$env:STRETCH_SESSION_PID)
        if (-not $ok) { $allGood = $false }
        Write-TestCase -Name 'designated-pid' -Pass:$ok -Fail:(-not $ok) `
                       -Detail "session=$($sx.ProcId) designated=$($env:STRETCH_SESSION_PID)"
    }
    else {
        Write-TestCase -Name 'designated-pid' -Pass -Detail 'standalone -- no designated pid to compare'
    }

    # 2. The bound Application reports the same main window we were given --
    #    the object model and the window agree about which Excel this is.
    $appHwnd = [int64]$sx.App.Hwnd
    $ok2 = ($appHwnd -eq [int64]$sx.Hwnd)
    if (-not $ok2) { $allGood = $false }
    Write-TestCase -Name 'app-hwnd-matches' -Pass:$ok2 -Fail:(-not $ok2) `
                   -Detail "app.Hwnd=$appHwnd session.Hwnd=$([int64]$sx.Hwnd)"

    # 3. A real object-model round trip THROUGH the binding: write a value
    #    into the baseline workbook, read it back. Proves the proxy is live,
    #    not a husk that answered two property reads from cache.
    #
    #    EVERY LINK IN THE CHAIN IS HELD AND RELEASED. The first version of
    #    this test wrote it as one dotted chain, which leaks an intermediate
    #    RCW per dot -- and its session sat as a refcount zombie for the full
    #    close deadline, dump on file. This is the pattern a well-behaved
    #    test follows (Complete-Test also sweeps stragglers, but a test that
    #    relies on the sweep is leaning on the safety net, not the contract).
    $booksRef = $sx.App.Workbooks
    $bookRef  = $booksRef.Item(1)
    $sheetsRef = $bookRef.Worksheets
    $sheetRef  = $sheetsRef.Item(1)
    $cellRef   = $sheetRef.Range('A1')
    $cellRef.Value2 = 42137
    $readBack = [int]$cellRef.Value2
    $cellRef.Value2 = $null                       # leave the baseline as found
    # THE WRITE DIRTIED THE WORKBOOK, and a dirty workbook turns the session
    # close into a hidden "Save changes?" dialog that waits forever -- this
    # test hung its session for the full deadline until this line existed
    # (dump on file). A test that modifies the baseline restores Saved unless
    # save behaviour is the thing it is testing.
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
