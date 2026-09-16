# ByRef (VBA's default) for every type; each has its own load opcode.
# Each parameter is read (to emit the load) and written (so ByRef is not demoted).
$case = @{ Name='frame-types-byref'
     Setup=@'
Public Sub T_RfByte(ByRef v As Byte)
    Dim z As Byte
    z = v
    v = z
End Sub
Public Sub T_RfInt(ByRef v As Integer)
    Dim z As Integer
    z = v
    v = z
End Sub
Public Sub T_RfLong(ByRef v As Long)
    Dim z As Long
    z = v
    v = z
End Sub
Public Sub T_RfLL(ByRef v As LongLong)
    Dim z As LongLong
    z = v
    v = z
End Sub
Public Sub T_RfSng(ByRef v As Single)
    Dim z As Single
    z = v
    v = z
End Sub
Public Sub T_RfDbl(ByRef v As Double)
    Dim z As Double
    z = v
    v = z
End Sub
Public Sub T_RfCy(ByRef v As Currency)
    Dim z As Currency
    z = v
    v = z
End Sub
Public Sub T_RfDate(ByRef v As Date)
    Dim z As Date
    z = v
    v = z
End Sub
Public Sub T_RfStr(ByRef v As String)
    Dim z As String
    z = v
    v = z
End Sub
Public Sub T_RfBool(ByRef v As Boolean)
    Dim z As Boolean
    z = v
    v = z
End Sub
Public Sub T_RfVar(ByRef v As Variant)
    Dim z As Variant
    z = v
    v = z
End Sub
Public Sub T_RfObj(ByRef v As Object)
    Dim z As Object
    Set z = v
    Set v = z
End Sub
Public Sub T_RfPtr(ByRef v As LongPtr)
    Dim z As LongPtr
    z = v
    v = z
End Sub
Public Sub T_RfDrive()
    Dim vByte As Byte
    Dim vInt As Integer
    Dim vLong As Long
    Dim vLL As LongLong
    Dim vSng As Single
    Dim vDbl As Double
    Dim vCy As Currency
    Dim vDate As Date
    Dim vStr As String
    Dim vBool As Boolean
    Dim vVar As Variant
    Dim vObj As Object
    Dim vPtr As LongPtr
    vByte = 200
    vInt = 1234
    vLong = &H11223344
    vLL = 1234567890123^
    vSng = 1.5
    vDbl = 2748.5
    vCy = 9.99
    vDate = #1/2/2020#
    vStr = "REFSTR"
    vBool = True
    vVar = 42
    Set vObj = Application
    vPtr = 4096
    Call T_RfByte(vByte)
    Call T_RfInt(vInt)
    Call T_RfLong(vLong)
    Call T_RfLL(vLL)
    Call T_RfSng(vSng)
    Call T_RfDbl(vDbl)
    Call T_RfCy(vCy)
    Call T_RfDate(vDate)
    Call T_RfStr(vStr)
    Call T_RfBool(vBool)
    Call T_RfVar(vVar)
    Call T_RfObj(vObj)
    Call T_RfPtr(vPtr)
End Sub
'@
     Invoke=@{ Name='T_RfDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.procedures -lt 14) {
            return "expected 14 procedures, got $($t.procedures)" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        # Every one of these parameters IS read, so every one has a load opcode
        # in the p-code. A '?' here is a missing entry in the type table, not an
        # absence of information -- and the message names which, so the gap is
        # actionable rather than merely reported.
        $missing = @()
        foreach ($fn in @('T_RfByte','T_RfInt','T_RfLong','T_RfLL','T_RfSng',
                          'T_RfDbl','T_RfCy','T_RfDate','T_RfStr','T_RfBool',
                          'T_RfVar','T_RfObj','T_RfPtr')) {
            if (-not $entry.ContainsKey($fn)) { $missing += "$fn=MISSING"; continue }
            if ($entry[$fn].typetext -match '\?') {
                $missing += "$fn$($entry[$fn].typetext)" }
        }
        if ($missing.Count -gt 0) {
            return "no recoverable type for: $($missing -join ' ')" }
        $null }
     Why='ByRef is VBA''s default and each type has its own load opcode; only
          Long& and String& had ever been measured' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
