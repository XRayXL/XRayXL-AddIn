# A parameter that is only forwarded must not claim a type: opcode 751, the generic by-reference
# push, says nothing about the slot, so it must not read `Udt&`. A real UDT is the control.
$case = @{ Name='params-only-passed-on'
     Setup=@'
Private Type TPoint
    x As Double
    n As Long
End Type

' The callee, so the parameters above it have somewhere to be forwarded TO.
Private Sub P_Sink(ByRef n As Long)
    n = n + 1
End Sub
Private Sub P_SinkArr(ByRef a() As Double)
    a(0) = 1#
End Sub

' ONLY passed on: no read, no write, nothing that could name the type.
Public Sub P_FwdLong(ByRef n As Long)
    P_Sink n
End Sub
Public Sub P_FwdArr(ByRef a() As Double)
    P_SinkArr a
End Sub

' THE CONTROL: a genuine UDT reference, whose member access does name it.
Public Sub P_RealUdt(ByRef p As TPoint)
    p.n = p.n + 1
End Sub

Public Sub P_FwdDrive()
    Dim n As Long, a(0 To 3) As Double, p As TPoint
    n = 5
    P_FwdLong n
    P_FwdArr a
    P_RealUdt p
End Sub
'@
     Invoke=@{ Name='P_FwdDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        foreach ($fn in @('P_FwdLong','P_FwdArr','P_RealUdt')) {
            if (-not $entry.ContainsKey($fn)) { return "never traced: $fn" }
        }
        # an unresolved marker is the right answer: the opcode does not carry the type
        $wrong = @()
        foreach ($fn in @('P_FwdLong','P_FwdArr')) {
            $tt = [string]$entry[$fn].typetext
            if ($tt -match 'Udt') { $wrong += "$fn=$tt" }
        }
        if ($wrong.Count -gt 0) {
            return "a forwarded parameter claimed a UDT it is not: $($wrong -join ' ')" }
        # a real UDT is named through its member access, a different opcode family
        $udt = [string]$entry['P_RealUdt'].typetext
        if ($udt -notmatch 'Udt') {
            return "the real UDT stopped being recognised: P_RealUdt=$udt" }
        $null }
     Why='751 is the generic by-reference push, not a UDT marker; it was tabled
          as `Udt&` from a single measurement whose UDT was incidental, so a
          forwarded `ByRef Long` reported a confident wrong type' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
