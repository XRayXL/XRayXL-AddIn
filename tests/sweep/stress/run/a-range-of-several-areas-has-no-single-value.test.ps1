$case = @{ Name='a-range-of-several-areas-has-no-single-value'
     Modules=@{
       'M'=@'
' Value2 of `A1:A2,C1:C2` is the first area's alone, so a described Range of several areas
' carries its address and no values; one area carries both.
Public Function Take(ByVal r As Variant) As Long
    Take = r.Count
End Function

Public Sub Drive()
    Dim sh As Object
    Set sh = Application.ThisWorkbook.Worksheets("S1")
    sh.Range("E1").Value = Take(sh.Range("A1:A2,C1:C2"))
    sh.Range("E2").Value = Take(sh.Range("A1:A2"))
End Sub
'@
     }
     Cells=@{ 'A1' = '11'; 'A2' = '12'; 'C1' = '31'; 'C2' = '32' }
     Trigger=@{ Kind='Run'; Name='Drive' }
     Expect={ param($t)
        $takes = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq 'Take' })
        if ($takes.Count -ne 2) { return "Take was traced $($takes.Count) time(s), expected 2" }
        if ([string]$t.cells['E1'] -ne '4') { return "VBA counted $($t.cells['E1']) cells in the two areas" }
        $two = ArgOf $takes[0] 1
        $one = ArgOf $takes[1] 1
        if ($two -notmatch '^Range@0x[0-9A-F]+\(.*A1:A2,.*C1:C2\)$') {
            return "the two-area Range reads [$two]; expected its address and no values" }
        if ($one -notmatch '^Range@0x[0-9A-F]+\(.*A1:A2\)=.*11.*12') {
            return "the one-area Range reads [$one]; expected its address and its values" }
        $null }
     Why='Range.Value2 of several areas answers for the first alone, so the values
          would describe part of the range as if it were the whole' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
