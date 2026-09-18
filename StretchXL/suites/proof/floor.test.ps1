# PROOF: the simplest possible test. Connects to whatever Excel the contract
# provides, checks the session baseline (one workbook exists), and passes.
# If this fails, the fault is the harness or the kit, never a workload --
# which is exactly what makes it worth running first in every suite tree.
. (Join-Path $PSScriptRoot '..\..\TestKit.ps1')

try {
    $sx = Connect-TestExcel
    # Held and released, not a dotted chain, which leaks an RCW per dot (binding.test.ps1).
    $booksRef = $sx.App.Workbooks
    $bookCount = $booksRef.Count
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($booksRef)
    if ($bookCount -ge 1) {
        Complete-Test -Pass -Detail "managed=$($sx.Managed) books=$bookCount"
    }
    else {
        Complete-Test -Fail -Detail "expected the baseline workbook, found $bookCount"
    }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
