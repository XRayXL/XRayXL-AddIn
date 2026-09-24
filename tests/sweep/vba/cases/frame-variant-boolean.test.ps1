# A Boolean inside a Variant or an array renders TRUE/FALSE (a plain one reads as an Integer).
# False is planted, since a renderer writing TRUE for both would pass on True alone.
$case = @{ Name='frame-variant-boolean'
     Setup=@'
Public Sub T_VbFalse(ByVal v As Variant)
    Dim z As Variant
    z = v
End Sub
Public Sub T_VbTrue(ByVal v As Variant)
    Dim z As Variant
    z = v
End Sub
Public Sub T_VbArr(ByRef v() As Boolean)
    Dim z As Boolean
    z = v(0)
End Sub
Public Sub T_VbDrive()
    Dim b(1) As Boolean
    b(0) = False
    b(1) = True
    Call T_VbFalse(False)
    Call T_VbTrue(True)
    Call T_VbArr(b)
End Sub
'@
     Invoke=@{ Name='T_VbDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        $e = Get-FirstEntryByName $t.rows
        foreach ($fn in 'T_VbFalse', 'T_VbTrue', 'T_VbArr') {
            if (-not $e[$fn]) { return "$fn was not traced" } }
        $f = [string]$e['T_VbFalse'].args
        if ($f -notmatch 'FALSE' -or $f -match 'TRUE') { return "T_VbFalse args were [$f], expected FALSE and no TRUE" }
        $tr = [string]$e['T_VbTrue'].args
        if ($tr -notmatch 'TRUE' -or $tr -match 'FALSE') { return "T_VbTrue args were [$tr], expected TRUE and no FALSE" }
        $a = [string]$e['T_VbArr'].args
        if ($a -notmatch 'FALSE' -or $a -notmatch 'TRUE') { return "T_VbArr args were [$a], expected both FALSE and TRUE" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     Why='FALSE and TRUE render differently inside a Variant and an array; only True was ever planted' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
