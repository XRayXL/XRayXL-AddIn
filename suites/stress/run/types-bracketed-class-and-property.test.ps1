# Declaration forms nothing else reaches: `As [Long]` (must decode as Long),
# a user class as a parameter type, and a Property Let's trailing value parameter.
$case = @{ Name='types-bracketed-class-and-property'
     Modules=@{
       'M'=@'
Public Sub BracketLong(ByVal v As [Long])
    Dim z As Long
    z = v
End Sub
Public Sub BracketVariant(ByVal v As [Variant])
    Dim z As Variant
    z = v
End Sub
Public Sub ClassParam(ByVal c As CThing2)
    Dim z As Long
    z = c.Val
End Sub
Public Sub Go()
    Dim c As CThing2
    Set c = New CThing2
    c.Val = &H31313131
    Call BracketLong(&H11223344)
    Call BracketVariant(42)
    Call ClassParam(c)
End Sub
'@
     }
     Classes=@{
       'CThing2'=@'
Private mV As Long
Public Property Let Val(ByVal v As Long)
    mV = v
End Property
Public Property Get Val() As Long
    Val = mV
End Property
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $e = Get-FirstEntryByName $t.rows
        # A bracketed reserved type name IS the type -- same opcode, same value.
        if (-not $e.ContainsKey('BracketLong')) { return 'BracketLong was not traced' }
        if ($e['BracketLong'].typetext -ne 'Long') {
            return "BracketLong signature was [$($e['BracketLong'].typetext)], expected Long -- As [Long] must decode as As Long" }
        if ($e['BracketLong'].args -ne 'a1:Long=287454020') {
            return "BracketLong args were [$($e['BracketLong'].args)]" }
        if ($e['BracketVariant'].typetext -ne 'Variant') {
            return "BracketVariant signature was [$($e['BracketVariant'].typetext)], expected Variant" }
        # A class-module type is an Object, as a library class is: the class
        # NAME is compile-time and is simply not present in the p-code.
        if ($e['ClassParam'].typetext -ne 'Object') {
            return "ClassParam signature was [$($e['ClassParam'].typetext)], expected Object" }
        # THE PROPERTY LET VALUE-PARAM, AND THE IMPLICIT `Me`.
        #
        # MS-VBAL: every method has an implicit ByVal `Me` (the current
        # object). If that took an argument slot, every class method's
        # arguments would be shifted by one and `a1` would be the object --
        # silently wrong on every class in a workbook. It does not: a class
        # Property Let reports ONE parameter holding the assigned value, so
        # declared parameters still start at slot 1 exactly as in a standard
        # module. Asserted exactly, because a lenient check here proves
        # nothing about the thing worth checking.
        if (-not $e.ContainsKey('Val')) { return 'Property Let Val was not traced' }
        if ($e['Val'].typetext -ne 'Long') {
            return "Property Val signature was [$($e['Val'].typetext)], expected Long -- an implicit Me would shift it" }
        if ($e['Val'].args -ne 'a1:Long=825307441') {
            return "Property Val args were [$($e['Val'].args)], expected a1:Long=825307441 (the assigned value, not Me)" }
        $null }
     Why='bracketed reserved type names, a user-defined class module as a
          parameter type, and the trailing value-param of a Property Let' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
