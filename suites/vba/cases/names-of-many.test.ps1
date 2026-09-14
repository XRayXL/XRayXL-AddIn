$case = @{ Name='names-of-many'
     Setup=@'
Public Sub T_NameA()
    Call T_NameB
    Call T_NameC
End Sub
Public Sub T_NameB()
    Call T_NameD
End Sub
Public Sub T_NameC()
    Dim a As Long
    a = 1
End Sub
Public Sub T_NameD()
    Dim a As Long
    a = 1
End Sub
'@
     Invoke=@{ Name='T_NameA'; Args=@() }
     Expect={ param($t)
        if ($t.unnamed -ne 0) { return "$($t.unnamed) unnamed procedure(s)" }
        if ($t.named -lt 4)   { return "expected 4 named procedures, got $($t.named)" }
        # THE WORKBOOK NAME, which nothing asserted until now.
        #
        # Module and function names are checked by 94 assertions across this
        # suite, so the offsets on THAT path are well covered. The workbook
        # comes down a different branch of the same chain --
        #   parent -> pObjTbl(0x08) -> pOwner(0x08) -> wszFilename(0x26)
        # -- and every one of those three offsets was unverified: a wrong one
        # attributes a procedure to the wrong workbook, which with several
        # open is exactly the question the tool exists to answer.
        $mods = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.module } |
                            ForEach-Object { $_.module } | Sort-Object -Unique)
        if (-not $mods) { return "no module/workbook recorded on any entry row" }
        foreach ($m in $mods) {
            if ($m -notmatch '^\[VbaRun\.xlsm\]') {
                return "workbook name wrong: got [$m], expected it to start [VbaRun.xlsm]" }
        }
        foreach ($want in @('T_NameA','T_NameB','T_NameC','T_NameD')) {
            if ($t.names -notcontains $want) {
                return "name '$want' missing; resolved: $($t.names -join ',')" }
        }
        $null }
     Why='four procedures in one module, each reached by a different path --
          the procMap index has to be right for every one, not just the first' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
