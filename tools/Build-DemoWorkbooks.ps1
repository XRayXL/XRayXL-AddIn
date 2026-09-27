<#
  Regenerates dist\demo\01..04 and 06..10 *.xlsm; 05_VBACurves.xlsm is hand-built.

    .\tools\Build-DemoWorkbooks.ps1              every workbook
    .\tools\Build-DemoWorkbooks.ps1 -Only 10     just 10_Plasma.xlsm; the others are left as they are

  Driving Excel to inject VBA needs "Trust access to the VBA project object model",
  so the .xlsm files are committed and this runs only when the demo content changes.
  XRayXL is driven from its ribbon, so a button appears only where the traced thing
  is itself a macro.
#>
param(
    # workbook numbers to save, such as 10; the rest are built but not saved
    [string[]]$Only = @()
)
$ErrorActionPreference = 'Stop'

$root  = Split-Path $PSScriptRoot -Parent
$wbOut = Join-Path $root 'dist\demo'
$built = Join-Path $root 'build\x64\Release'
New-Item -ItemType Directory -Force $wbOut | Out-Null

# From the build output, never dist\, which holds a release. Build them first
# (msbuild XRayXL.sln) or the workbooks save with #NAME? cached in every cell.
$registerXlls = @(
    (Join-Path $built 'DemoFinance\DemoFinance64.xll')
    (Join-Path $built 'DemoBehaviors\DemoBehaviors64.xll')
)

# superseded workbook names, deleted if still present
$retired = @('01_CalcChain.xlsm', '02_Events.xlsm', '03_Errors.xlsm', '04_Advanced.xlsm')

function Clear-ComRef($Refs) {
    # an unreleased COM wrapper keeps Excel alive after Quit
    foreach ($r in @($Refs)) {
        if ($r) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($r) } catch {} }
    }
}

Add-Type -Namespace DemoBuild -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
'@

$xl = New-Object -ComObject Excel.Application
[uint32]$xlPid = 0
[void][DemoBuild.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
$xl.Visible = $false
$xl.DisplayAlerts = $false
try { $xl.EnableEvents = $false } catch {}   # no events while building

# try/finally, so a failed build does not leave a hidden Excel running
try {
    foreach ($p in $registerXlls) {
        if (Test-Path $p) { try { [void]$xl.RegisterXLL($p) } catch { Write-Warning "RegisterXLL $(Split-Path $p -Leaf): $($_.Exception.Message)" } }
        else { Write-Warning "not built, so its formulas will cache as #NAME?: $p" }
    }

    # Prove VBA trust once, up front, with a clear message. The finally below quits Excel.
    try { $probe = $xl.Workbooks.Add(); [void]$probe.VBProject; $probe.Close($false) }
    catch { throw "Cannot inject VBA: enable File > Options > Trust Center > Macro Settings > 'Trust access to the VBA project object model', then re-run." }
    finally { Clear-ComRef @($probe) }

    function New-DemoBook([string]$sheetName) {
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = $sheetName
        $xl.ActiveWindow.DisplayGridlines = $false
        $ws.Cells.Font.Name = 'Segoe UI'
        $ws.Cells.Font.Size = 10
        return @{ Book = $wb; Sheet = $ws; Project = $wb.VBProject }
    }

    function Add-Module($proj, $name, $code) {
        $c = $proj.VBComponents.Add(1); $c.Name = $name       # 1 = standard module
        $c.CodeModule.AddFromString($code)
        Clear-ComRef @($c)
    }

    function Set-SheetCode($proj, $codeName, $code) {
        $comp = $proj.VBComponents.Item($codeName)
        $comp.CodeModule.AddFromString($code)
        Clear-ComRef @($comp)
    }

    # The top of every sheet: a title, one line on what it shows, and the steps to run it.
    # Returns the first free row.
    function Write-Guide($ws, [string]$title, [string]$purpose, [string[]]$steps) {
        $t = $ws.Range('A1'); $t.Value2 = $title; $t.Font.Size = 16; $t.Font.Bold = $true
        $ws.Range('A2').Value2 = $purpose
        $ws.Range('A2').Font.Italic = $true
        $r = 4
        $h = $ws.Range("A$r"); $h.Value2 = 'Try it'; $h.Font.Bold = $true
        $r++
        $n = 1
        foreach ($s in $steps) { $ws.Range("A$r").Value2 = "$n.  $s"; $r++; $n++ }
        return $r + 1
    }

    # A section heading with a rule under it.
    function Write-Heading($ws, [int]$row, [string]$text) {
        $c = $ws.Range("A$row"); $c.Value2 = $text; $c.Font.Bold = $true; $c.Font.Size = 11
        $ws.Range("A${row}:H${row}").Borders.Item(9).LineStyle = 1   # 9 = xlEdgeBottom
    }

    # One demo row: what it is, the formula, and what to look for in the trace.
    function Write-Demo($ws, [int]$row, [string]$label, [string]$formula, [string]$look, [switch]$Spill) {
        $ws.Range("A$row").Value2 = $label
        if ($Spill) { $ws.Range("B$row").Formula2 = $formula } else { $ws.Range("B$row").Formula = $formula }
        $ws.Range("B$row").HorizontalAlignment = -4131   # xlLeft
        if ($look) { $ws.Range("J$row").Value2 = $look; $ws.Range("J$row").Font.Color = 0x606060 }
    }

    function Add-Button($ws, [string]$name, [string]$text, [string]$macro, $anchor) {
        $cell = $ws.Range($anchor)
        $b = $ws.Buttons().Add($cell.Left, $cell.Top, 150, 22)
        $b.Name = $name; $b.Text = $text; $b.OnAction = $macro
        Clear-ComRef @($b, $cell)
    }

    # Saves over dist\demo\<leaf>, unless -Only leaves it out.
    function Save-As($demo, [string]$leaf) {
        $demo.Sheet.Range('A1').Select() | Out-Null
        if ($Only.Count -eq 0 -or ($Only | Where-Object { $leaf.StartsWith($_) })) {
            $path = Join-Path $wbOut $leaf
            Remove-Item $path -ErrorAction SilentlyContinue
            $demo.Book.SaveAs($path, 52)      # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
            Write-Host "  workbook -> $leaf"
        }
        $demo.Book.Close($false)
        Clear-ComRef @($demo.Project, $demo.Sheet, $demo.Book)
    }

    function Save-Book($demo, [string]$leaf) {
        $ws = $demo.Sheet
        $ws.Columns('A').ColumnWidth = 44
        $ws.Columns('B:I').ColumnWidth = 11
        $ws.Columns('J').ColumnWidth = 70
        Save-As $demo $leaf
    }

    $armStep     = 'Developer tab > XRayXL > Arm.'
    $disarmStep  = 'Developer tab > XRayXL > Disarm.'
    $readStep    = 'Open the trace: Options > Output > right-click Trace file > Reveal in File Explorer. (Tail on the ribbon, pressed after arming, shows rows as they arrive.)'

    # ===== 01 -- the first trace: VBA and XLL in one recalculation ============
    $d = New-DemoBook 'First trace'
    Add-Module $d.Project 'FirstTrace' @'
' A VBA function that calls a nested VBA helper: one cell, two VBA frames.
Public Function RiskWeighted(ByVal exposure As Double, ByVal rating As Long) As Double
    RiskWeighted = exposure * WeightFor(rating)
End Function

Private Function WeightFor(ByVal rating As Long) As Double
    Select Case rating
        Case 1: WeightFor = 0.2
        Case 2: WeightFor = 0.5
        Case Else: WeightFor = 1#
    End Select
End Function

' A VBA function that calls an XLL function: the XLL row nests inside the VBA one.
Public Function OptionBook(ByVal spot As Double) As Double
    OptionBook = 100# * Application.Run("BlackScholes", spot, 100#, 1#, 0.05, 0.2)
End Function
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '01  Your first trace' `
        'One recalculation traced: VBA functions, XLL functions, and one calling the other.' `
        @($armStep, 'Press Ctrl+Alt+F9 to recalculate everything.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What gets calculated'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    Write-Demo $ws $r 'RiskWeighted -- VBA calling VBA' '=RiskWeighted(1000000,3)' `
        'RiskWeighted at depth 1, then WeightFor at depth 2 with parent = RiskWeighted''s span'; $r++
    Write-Demo $ws $r 'OptionBook -- VBA calling an XLL' '=OptionBook(100)' `
        'BlackScholes''s rows fall between OptionBook''s entry and exit, same callerref; each source counts its own depth'; $r++
    Write-Demo $ws $r 'BlackScholes -- an XLL function' '=BlackScholes(100,100,1,0.05,0.2)' `
        'source XLL; args a1:B=100 ... ; ret and rettype Q'; $r++
    Write-Demo $ws $r 'SlowSum -- a slow XLL function' '=SlowSum(2,3)' `
        'the ticks column: this one is far slower than its neighbours'; $r++
    Write-Demo $ws $r 'Fibonacci -- a heavy XLL function' '=Fibonacci(30)' `
        'ticks again; callerref names the cell each call came from'; $r++
    Save-Book $d '01_FirstTrace.xlsm'

    # ===== 02 -- values: ranges, arrays, Variants and types =====================
    $d = New-DemoBook 'Values'
    Add-Module $d.Project 'Values' @'
' A range argument: the trace shows the Range, its address, and its cells row by row.
Public Function CellCount(ByVal cells As Variant) As Long
    CellCount = cells.Count
End Function

' A Variant array of mixed types: in a Variant, a number other than a Double names its type.
Public Function MixedTypes() As Variant
    MixedTypes = Array(CInt(1), 2.5, CLng(3), CCur(1.5), DateSerial(2026, 9, 19), "text", True)
End Function

' A 2-D typed array: both bounds of each dimension, then one brace level per row.
Public Function TimesTable(ByVal n As Long) As Variant
    Dim t() As Long, r As Long, c As Long
    ReDim t(1 To n, 1 To n)
    For r = 1 To n
        For c = 1 To n
            t(r, c) = r * c
        Next c
    Next r
    TimesTable = t
End Function

' An array handed on by reference: the bytecode says only "a reference", so the trace
' recognises the array from its own descriptor rather than from a declaration.
Public Function SumOfSquares(ByVal n As Long) As Double
    Dim a() As Double, i As Long
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = i * i
    Next i
    SumOfSquares = Total(a)
End Function

Private Function Total(ByRef squares() As Double) As Double
    Dim i As Long, s As Double
    For i = LBound(squares) To UBound(squares)
        s = s + squares(i)
    Next i
    Total = s
End Function
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '02  Arguments and values' `
        'How the trace writes what goes in and comes out: ranges, arrays, Variants and their types.' `
        @($armStep, 'Press Ctrl+Alt+F9 to recalculate everything.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What gets calculated'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    $dataRow = $r + 12
    Write-Demo $ws $r 'CellCount -- a range as an argument' "=CellCount(B${dataRow}:D$($dataRow + 2))" `
        "a1 is Range@0x...([02_Values.xlsm]Values!B${dataRow}:D$($dataRow + 2))= followed by Variant[1..3,1..3]{{...},{...},{...}}: row 1 first, each row in its own braces"; $r++
    Write-Demo $ws $r 'MixedTypes -- a Variant array of mixed types' '=MixedTypes()' `
        'ret Variant[0..6]{Integer(1),2.5,Long(3),Currency(1.5000),Date(46284),"text",TRUE}: a bare number is a Double' -Spill; $r++
    Write-Demo $ws $r 'TimesTable -- a 2-D typed array' '=TimesTable(3)' `
        'ret Long[1..3,1..3]{{1,2,3},{2,4,6},{3,6,9}}: the elements are bare because the header says Long' -Spill; $r += 3
    Write-Demo $ws $r 'SumOfSquares -- an array passed by reference' '=SumOfSquares(5)' `
        'Total''s a1 is Ref&=Double[1..5]{1,4,9,16,25}: recognised from the array itself (see Declared or inferred)'; $r++
    Write-Demo $ws $r 'MakeSeries -- an XLL returning an array' '=MakeSeries(4)' `
        'ret Variant[1..1,1..4]{{1,4,9,16}}: an XLL array reads like a VBA one' -Spill; $r += 5
    $ws.Range("A$dataRow").Value2 = 'The range CellCount reads:'
    $ws.Range("B$dataRow").Value2 = 1.5;         $ws.Range("C$dataRow").Value2 = 'text'; $ws.Range("D$dataRow").Value2 = $true
    $ws.Range("B$($dataRow + 1)").Value2 = 42;   $ws.Range("C$($dataRow + 1)").Formula = '=NA()'
    $ws.Range("B$($dataRow + 2)").Value2 = -2.25; $ws.Range("C$($dataRow + 2)").Value2 = 'more'; $ws.Range("D$($dataRow + 2)").Value2 = 46284
    $ws.Range("B${dataRow}:D$($dataRow + 2)").Borders.LineStyle = 1
    $ws.Range("J$dataRow").Value2 = 'numbers bare, text quoted, TRUE, #N/A, Empty for the blank cell'
    $ws.Range("J$dataRow").Font.Color = 0x606060
    Save-Book $d '02_Values.xlsm'

    # ===== 03 -- callers: what started each call ================================
    $d = New-DemoBook 'Callers'
    Add-Module $d.Project 'Callers' @'
' A plain VBA function in a cell: the caller is that cell.
Public Function Twice(ByVal x As Double) As Double
    Twice = 2 * x
End Function

' An array formula entered over three cells: the caller is the whole range.
Public Function ThreeOf(ByVal x As Double) As Variant
    ThreeOf = Array(x, x * 2, x * 3)
End Function

' Run by the button: the caller is the button.
Public Sub FromTheButton()
    Helper "button"
End Sub

' The button schedules this with Application.OnTime: Excel runs it a second later.
Public Sub StartTimer()
    Application.OnTime Now + TimeSerial(0, 0, 1), "TimerTick"
End Sub

Public Sub TimerTick()
    Helper "timer"
End Sub

Public Sub Helper(ByVal who As String)
    Dim note As String
    note = "called from the " & who
End Sub
'@
    Set-SheetCode $d.Project $d.Sheet.CodeName @'
' Runs when you edit the Trigger cell: VBA started by an event, not by a calculation.
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("Trigger")) Is Nothing Then Exit Sub
    Helper "change event"
End Sub
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '03  Who called it' `
        'The caller and callerref columns: a cell, an array formula, an event, a button and a timer.' `
        @($armStep, 'Press Ctrl+Alt+F9, type a number into the Trigger cell, press Run a macro, then press Start a timer and wait a second.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What starts a call'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    Write-Demo $ws $r 'Twice -- a function in a cell' '=Twice(7)' `
        "caller cell, callerref '[03_Callers.xlsm]Callers'!B$r"; $r++
    $ws.Range("A$r").Value2 = 'ThreeOf -- an array formula over three cells'
    $ws.Range("B${r}:D$r").FormulaArray = '=ThreeOf(5)'
    $ws.Range("J$r").Value2 = "caller cell, callerref the whole range ...!B${r}:D$r"; $ws.Range("J$r").Font.Color = 0x606060
    $r++
    Write-Demo $ws $r 'CallCounter -- a volatile XLL function' '=CallCounter()' `
        'volatile: it runs again whenever anything recalculates, as the Trigger edit shows'; $r++
    $ws.Range("A$r").Value2 = 'Trigger -- type a number here:'
    $ws.Range("B$r").Name = 'Trigger'; $ws.Range("B$r").Value2 = 0
    $ws.Range("B$r").Interior.Color = 0xCCF2FF
    $ws.Range("J$r").Value2 = 'Worksheet_Change, then Helper beneath it: VBA run by an event'; $ws.Range("J$r").Font.Color = 0x606060
    $r++
    $ws.Range("A$r").Value2 = 'A button that runs a macro:'
    Add-Button $ws 'RunAMacro' 'Run a macro' 'FromTheButton' "B$r"
    $ws.Range("J$r").Value2 = 'FromTheButton with caller name, callerref RunAMacro -- the button''s name'; $ws.Range("J$r").Font.Color = 0x606060
    $r++
    $ws.Range("A$r").Value2 = 'A timer (Application.OnTime):'
    Add-Button $ws 'StartATimer' 'Start a timer' 'StartTimer' "B$r"
    $ws.Range("J$r").Value2 = 'StartTimer from the button, then TimerTick a second later, run by Excel itself'; $ws.Range("J$r").Font.Color = 0x606060
    $r++
    foreach ($row in 1..$r) { $ws.Rows($row).RowHeight = 22 }
    Save-Book $d '03_Callers.xlsm'

    # ===== 04 -- errors: how each call ended ======================================
    $d = New-DemoBook 'Errors'
    Add-Module $d.Project 'Errors' @'
' Divides, and VBA raises "Division by zero" when b is 0.
Private Function Divide(ByVal a As Double, ByVal b As Double) As Double
    Divide = a / b
End Function

' Catches the error itself, so the cell shows "n/a": Divide threw, SafeRatio handled it.
Public Function SafeRatio(ByVal a As Double, ByVal b As Double) As Variant
    On Error GoTo Failed
    SafeRatio = Divide(a, b)
    Exit Function
Failed:
    SafeRatio = "n/a"
End Function

' Nothing catches it, so Excel shows #VALUE!: Divide threw, RawRatio did not handle it.
Public Function RawRatio(ByVal a As Double, ByVal b As Double) As Double
    RawRatio = Divide(a, b)
End Function

' Returning an error VALUE is not an error: the call returns normally, with #N/A.
Public Function CodeFor(ByVal ccy As String) As Variant
    If ccy = "EUR" Then CodeFor = 978 Else CodeFor = CVErr(xlErrNA)
End Function

' An Optional argument left out arrives as Missing.
Public Function Scaled(ByVal x As Double, Optional ByVal factor As Variant) As Double
    If IsMissing(factor) Then Scaled = x Else Scaled = x * factor
End Function

' Run by the button: Thrower raises, Middle passes it on, Outer catches it.
Public Sub RunErrorChain()
    Outer
End Sub

Private Sub Outer()
    On Error GoTo Caught
    Middle
    Exit Sub
Caught:
    Dim resumed As Long
    resumed = 1
End Sub

Private Sub Middle()
    Thrower
End Sub

Private Sub Thrower()
    Err.Raise 5, "Demo", "a deliberate error"
End Sub
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '04  Errors and how calls end' `
        'The outcome column: returned, threw, unwound, handled and unhandled -- and error values that are not errors.' `
        @($armStep, 'Press Ctrl+Alt+F9, then press Run the error chain.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What gets calculated'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    Write-Demo $ws $r 'SafeRatio(1,0) -- an error caught in VBA' '=SafeRatio(1,0)' `
        'Divide: outcome threw.  SafeRatio: outcome handled, ret "n/a"'; $r++
    Write-Demo $ws $r 'RawRatio(1,0) -- an error nothing catches' '=RawRatio(1,0)' `
        'Divide: threw.  RawRatio: unhandled, and Excel shows #VALUE!'; $r++
    Write-Demo $ws $r 'CodeFor("GBP") -- returning an error value' '=CodeFor("GBP")' `
        'outcome returned, ret #N/A: an error value is a result, not an error'; $r++
    Write-Demo $ws $r 'Scaled(10) -- an Optional left out' '=Scaled(10)' `
        'a2:Variant=Missing'; $r++
    Write-Demo $ws $r 'MightDivide(10,0) -- an XLL returning an error' '=MightDivide(10,0)' `
        'source XLL, ret #DIV/0!, outcome returned'; $r++
    Write-Demo $ws $r 'SafeDivide(10,0) -- an XLL that avoids it' '=SafeDivide(10,0)' `
        'the same inputs, a plain value back'; $r++
    $r++
    Write-Heading $ws $r 'An error passed up a chain of macros'
    $r++
    $ws.Range("A$r").Value2 = 'Outer calls Middle calls Thrower:'
    Add-Button $ws 'RunTheErrorChain' 'Run the error chain' 'RunErrorChain' "B$r"
    $ws.Range("J$r").Value2 = 'Thrower: threw.  Middle: unwound -- it did not catch it.  Outer: handled.'; $ws.Range("J$r").Font.Color = 0x606060
    foreach ($row in 1..$r) { $ws.Rows($row).RowHeight = 22 }
    Save-Book $d '04_Errors.xlsm'

    # ===== 06 -- objects: a class module, and Excel objects as arguments =========
    $d = New-DemoBook 'Objects'
    $cls = $d.Project.VBComponents.Add(2); $cls.Name = 'Position'     # 2 = class module
    $cls.CodeModule.AddFromString(@'
Private mQty As Double

Private Sub Class_Initialize()
    mQty = 0
End Sub

Public Property Let Qty(ByVal q As Double)
    mQty = q
End Property

Public Property Get Qty() As Double
    Qty = mQty
End Property

Public Function Value(ByVal px As Double) As Double
    Value = mQty * px
End Function

Private Sub Class_Terminate()
    mQty = 0
End Sub
'@)
    Clear-ComRef @($cls)
    Add-Module $d.Project 'Objects' @'
' Creates a Position, sets it up, values it, and lets it go: every class member shows as a call.
Public Function PositionValue(ByVal units As Double, ByVal px As Double) As Double
    Dim p As Position
    Set p = New Position
    p.Qty = units
    PositionValue = p.Value(px)
End Function

' A Range argument: the trace names it, gives its address and reads its cells.
Public Function RangeInfo(ByVal r As Range) As String
    RangeInfo = r.Address(False, False) & " has " & r.Cells.Count & " cells"
End Function

' Run by the button: a Worksheet, a Workbook and a Collection passed as arguments,
' and a Range returned.
Public Sub DescribeThings()
    Describe ActiveSheet
    Describe ThisWorkbook
    Dim items As New Collection
    items.Add "a": items.Add CLng(2)
    Describe items
    Dim c As Range
    Set c = FirstDataCell()
End Sub

Private Sub Describe(ByVal o As Object)
    Dim n As String
    n = TypeName(o)
End Sub

Private Function FirstDataCell() As Range
    Set FirstDataCell = ThisWorkbook.Worksheets("Objects").Range("DataStart")
End Function
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '06  Objects and classes' `
        'A class module''s members as calls, and Excel objects -- a Range, a Worksheet, a Workbook -- as arguments.' `
        @($armStep, 'Press Ctrl+Alt+F9, then press Describe the objects.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What gets calculated'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    $dataRow = $r + 6
    Write-Demo $ws $r 'PositionValue -- a class, created and released' '=PositionValue(100,2.5)' `
        'under PositionValue: Class_Initialize, Qty (a1:Double=100), Value (ret 250); then Class_Terminate, after PositionValue''s exit, when VBA releases the object'; $r++
    Write-Demo $ws $r 'RangeInfo -- a Range as an argument' "=RangeInfo(B${dataRow}:C$($dataRow + 1))" `
        "a1:Object=Range@0x...([06_Objects.xlsm]Objects!B${dataRow}:C$($dataRow + 1))=Variant[1..2,1..2]{{1.5,`"a`"},{2.5,`"b`"}}"; $r++
    $r++
    $ws.Range("A$r").Value2 = 'Objects passed by a macro:'
    Add-Button $ws 'DescribeTheObjects' 'Describe the objects' 'DescribeThings' "B$r"
    $ws.Range("J$r").Value2 = 'Describe three times: Worksheet@0x...([06_Objects.xlsm]Objects), Workbook@0x...([06_Objects.xlsm]), Collection@0x...=Variant[1..2]{"a",Long(2)}; FirstDataCell ret Range@0x...(...!B' + $dataRow + ')=1.5'
    $ws.Range("J$r").Font.Color = 0x606060
    $ws.Range("A$dataRow").Value2 = 'The range RangeInfo reads:'
    $ws.Range("B$dataRow").Name = 'DataStart'
    $ws.Range("B$dataRow").Value2 = 1.5;       $ws.Range("C$dataRow").Value2 = 'a'
    $ws.Range("B$($dataRow + 1)").Value2 = 2.5; $ws.Range("C$($dataRow + 1)").Value2 = 'b'
    $ws.Range("B${dataRow}:C$($dataRow + 1)").Borders.LineStyle = 1
    foreach ($row in 1..($dataRow + 1)) { $ws.Rows($row).RowHeight = 22 }
    Save-Book $d '06_Objects.xlsm'

    # ===== 07 -- the call tree: recursion, ByRef changes, and End ================
    $d = New-DemoBook 'Call tree'
    Add-Module $d.Project 'CallTree' @'
' Straight recursion: one chain, five deep.
Public Function Factorial(ByVal n As Long) As Double
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

' Recursion that branches: nine calls, and only `parent` says which called which.
Public Function VbaFib(ByVal n As Long) As Long
    If n < 2 Then
        VbaFib = n
    Else
        VbaFib = VbaFib(n - 1) + VbaFib(n - 2)
    End If
End Function

' ClampInPlace changes its ByRef argument, and the exit row shows the new value.
Public Function Clamped(ByVal x As Double) As Double
    Dim v As Double
    v = x
    ClampInPlace v, 0, 100
    Clamped = v
End Function

Private Sub ClampInPlace(ByRef v As Double, ByVal lo As Double, ByVal hi As Double)
    If v < lo Then v = lo
    If v > hi Then v = hi
End Sub

' Run by the button: End stops all VBA at once, and no procedure returns.
Public Sub StopEverything()
    Level1
End Sub

Private Sub Level1()
    Level2
End Sub

Private Sub Level2()
    End
End Sub
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '07  The call tree' `
        'depth and parent through recursion, a ByRef argument changing, and End stopping everything.' `
        @($armStep, 'Press Ctrl+Alt+F9, then press Stop everything.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What gets calculated'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    Write-Demo $ws $r 'Factorial(5) -- one chain, five deep' '=Factorial(5)' `
        'depth 1 to 5, each parent the span above it; a1 5,4,3,2,1; exits innermost first, ret 1,2,6,24,120'; $r++
    Write-Demo $ws $r 'VbaFib(4) -- recursion that branches' '=VbaFib(4)' `
        'nine VbaFib calls, deepest at depth 4; two calls at the same depth have different parents'; $r++
    Write-Demo $ws $r 'Clamped(150) -- a ByRef argument changed' '=Clamped(150)' `
        'ClampInPlace entry a1:Double&=150; its exit row carries a1:Double&=100, the new value'; $r++
    Write-Demo $ws $r 'Clamped(50) -- a ByRef argument left alone' '=Clamped(50)' `
        'ClampInPlace exit row carries no args: nothing changed'; $r++
    $r++
    $ws.Range("A$r").Value2 = 'End, three calls deep:'
    Add-Button $ws 'StopEverything' 'Stop everything' 'StopEverything' "B$r"
    $ws.Range("J$r").Value2 = 'StopEverything, Level1, Level2: outcome abandoned, trust end -- none of them returned'
    $ws.Range("J$r").Font.Color = 0x606060
    foreach ($row in 1..$r) { $ws.Rows($row).RowHeight = 22 }
    Save-Book $d '07_CallTree.xlsm'

    # ===== 08 -- threads: multi-threaded recalculation ============================
    $d = New-DemoBook 'Threads'
    Add-Module $d.Project 'Threads' @'
' VBA always runs on Excel's main thread, whatever the calculation settings.
Public Function VbaSquare(ByVal x As Double) As Double
    VbaSquare = x * x
End Function
'@
    $ws = $d.Sheet
    $r = Write-Guide $ws '08  Threads' `
        'The thread column: a thread-safe XLL spread across calculation threads, and VBA on the main thread.' `
        @('Check File > Options > Advanced > Formulas: Enable multi-threaded calculation.', $armStep,
          'Press Ctrl+Alt+F9.', $disarmStep, $readStep)
    Write-Heading $ws $r 'What gets calculated'
    $ws.Range("J$r").Value2 = 'Look for this in the trace'; $ws.Range("J$r").Font.Bold = $true
    $r++
    Write-Demo $ws $r 'ReverseText -- an XLL taking and returning text' '=ReverseText("hello")' `
        'typetext C%, a1:C%="hello", ret "olleh"'; $r++
    $r++
    $ws.Range("A$r").Value2 = 'Columns D and E, 500 rows each:'
    $ws.Range("J$r").Value2 = 'ThreadSafeSquare (rettype Q$) on several thread ids; VbaSquare all on one'
    $ws.Range("J$r").Font.Color = 0x606060
    $top = $r
    $ws.Range("D$top").Value2 = 'ThreadSafeSquare'; $ws.Range("E$top").Value2 = 'VbaSquare'
    $ws.Range("D${top}:E$top").Font.Bold = $true
    $ws.Range("D$($top + 1):D$($top + 500)").Formula = '=ThreadSafeSquare(ROW())'
    $ws.Range("E$($top + 1):E$($top + 500)").Formula = '=VbaSquare(ROW())'
    $ws.Columns('D:E').ColumnWidth = 16
    Save-Book $d '08_Threads.xlsm'

    # ===== 09 -- a bigger model, all XLL: an option book ==========================
    # Fixed rows, so the addresses on the sheet and in the walkthrough are the trace's own.
    $d = New-DemoBook 'Pricing'
    $ws = $d.Sheet
    [void](Write-Guide $ws '09  An option book, all XLL' `
        'A priced portfolio built from XLL functions: text, ranges, arrays and references as arguments, arrays as results.' `
        @($armStep, 'Press Ctrl+Alt+F9.', $disarmStep, $readStep))
    $ws.Range('A11').Value2 = 'Quotes';  $ws.Range('A11').Font.Bold = $true
    # A table goes in as one 2-D array, set through InvokeMember with every value unwrapped:
    # in this script PowerShell's COM binder takes Value2 for a string property and refuses
    # a number, and COM cannot marshal a PSObject wrapper.
    function Set-Table($ws, [string]$topLeft, $rows) {
        $n = $rows.Count; $m = $rows[0].Count
        $grid = [object[,]]::new($n, $m)
        for ($i = 0; $i -lt $n; $i++) { for ($j = 0; $j -lt $m; $j++) { $grid[$i, $j] = $rows[$i][$j].PSObject.BaseObject } }
        $target = $ws.Range($topLeft).Resize($n, $m)
        $argv = [object[]]::new(1); $argv[0] = $grid
        [void][System.__ComObject].InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $target, $argv)
    }
    Set-Table $ws 'A12' @(@('ACME', 102.5), @('BOLT', 48.2), @('CRUX', 250), @('DYNA', 75.4), @('EXPO', 12.9))
    $ws.Range('A12:B16').Name = 'Quotes'
    $ws.Range('D12').Value2 = 'Rate'; $ws.Range('E12').Value2 = 0.05; $ws.Range('E12').Name = 'Rate'

    $ws.Range('A18').Value2 = 'Positions'; $ws.Range('A18').Font.Bold = $true
    Set-Table $ws 'A19' @(,@('Ticker', 'Strike', 'Years', 'Vol', 'Qty', 'Spot', 'Price', 'Value'))
    $ws.Range('A19:H19').Font.Bold = $true
    Set-Table $ws 'A20' @(
        @('ACME', 100, 1, 0.20, 10), @('ACME', 110, 0.5, 0.25, -5), @('BOLT', 50, 2, 0.30, 20),
        @('BOLT', 45, 1, 0.30, 15),  @('CRUX', 240, 1, 0.18, 2),    @('CRUX', 260, 3, 0.22, 4),
        @('DYNA', 80, 0.25, 0.35, 8), @('DYNA', 70, 1, 0.35, -6),   @('EXPO', 12, 2, 0.40, 50),
        @('ZZZZ', 10, 1, 0.20, 1))
    $ws.Range('F20:F29').Formula = '=QuoteLookup(A20,Quotes)'
    $ws.Range('G20:G29').Formula = '=BlackScholes(F20,B20,C20,Rate,D20)'
    $ws.Range('H20:H29').Formula = '=E20*G20'
    $ws.Range('J20').Value2 = 'QuoteLookup ten times: a1:C%="ACME", a2:Q=Variant[1..5,1..2]{{"ACME",102.5},...}; ZZZZ returns #N/A'
    $ws.Range('J21').Value2 = 'BlackScholes nine times, not ten: Excel never calls it for row 29, whose spot is #N/A'

    $ws.Range('A31').Value2 = 'Portfolio'; $ws.Range('A31').Font.Bold = $true
    $ws.Range('A32').Value2 = 'Qty-weighted average price'; $ws.Range('B32').Formula = '=WeightedAverage(G20:G28,E20:E28)'
    $ws.Range('J32').Value2 = 'a1:K%=Double[1..9,1..1]{{...}} a2:K%=Double[1..9,1..1]{{10},{-5},...}: plain double arrays'
    $ws.Range('A33').Value2 = 'Cells in the positions table'; $ws.Range('B33').Formula = '=RangeSize(A20:H29)'
    $ws.Range('J33').Value2 = 'a1:U=SRef(R20C1:R29C8): a reference, named and not read; ret 80'
    $ws.Range('A34').Value2 = 'NPV of the bond cashflows'; $ws.Range('B34').Formula = '=NetPresentValue(Rate,B37:F37)'
    $ws.Range('J34').Value2 = 'a1:B=0.05 a2:Q=Variant[1..1,1..5]{{100,100,100,100,1100}}'
    $ws.Range('A35').Value2 = 'Grown for 10 years, then discounted'; $ws.Range('B35').Formula = '=PresentValue(CompoundReturn(1000,0.07,10),Rate,10)'
    $ws.Range('J35').Value2 = 'CompoundReturn then PresentValue, both depth 1 and callerref B35: the inner call finishes first'
    $ws.Range('A37').Value2 = 'Bond cashflows'
    Set-Table $ws 'B37' @(,@(100, 100, 100, 100, 1100))

    $ws.Range('A39').Value2 = 'Discount curve'; $ws.Range('A39').Font.Bold = $true
    Set-Table $ws 'B40' @(@(0.5), @(1), @(2), @(5), @(10))
    $ws.Range('C40').Formula2 = '=DiscountCurve(Rate,B40:B44)'
    $ws.Range('J40').Value2 = 'one call; ret Variant[1..5,1..1]{{0.97...},...}: a column in, a column out'

    $ws.Range('A46').Value2 = 'Price grid: strike across, vol down'; $ws.Range('A46').Font.Bold = $true
    Set-Table $ws 'B47' @(,@(80, 90, 100, 110, 120))
    Set-Table $ws 'A48' @(@(0.1), @(0.2), @(0.3), @(0.4))
    $ws.Range('B48:F51').Formula = '=BlackScholes(100,B$47,1,Rate,$A48)'
    $ws.Range('J48').Value2 = 'twenty more BlackScholes calls, one per cell, callerref B48 to F51'
    $ws.Range('J20:J48').Font.Color = 0x606060
    $ws.Columns('A').ColumnWidth = 34
    $ws.Columns('B:H').ColumnWidth = 11
    $ws.Columns('J').ColumnWidth = 90
    Save-As $d '09_XLLPricing.xlsm'

    # ===== 10 -- a plasma, painted two ways, for Perfetto =========================
    $d = New-DemoBook 'Plasma'
    Add-Module $d.Project 'Plasma' @'
Option Explicit

' A plasma, painted two ways. Act 1 sets each cell's colour, one call to Excel per cell. Act 2
' writes every cell's value in one call, and the sheet's colour scale paints them. Same maths.

Private Const GRID_ROWS As Long = 54
Private Const GRID_COLS As Long = 96
Private Const ACT_SECONDS As Double = 10

' The colour scale's three colours, so both acts look alike.
Private Const DEEP As Long = &H501014        ' RGB(20, 16, 80)
Private Const PINK As Long = &H8C32D6        ' RGB(214, 50, 140)
Private Const GOLD As Long = &H5AD6FF        ' RGB(255, 214, 90)

' Shared, not passed: an argument this size would be copied into every traced call's row.
Private field(1 To GRID_ROWS, 1 To GRID_COLS) As Double
Private grid As Range

Private running As Boolean, stopping As Boolean
Private showStart As Double, actStart As Double, actEnd As Double, lastTick As Double
Private actTitle As String, actFrames As Long

' The Start button.
Public Sub StartShow()
    If running Then Exit Sub
    running = True
    stopping = False
    On Error GoTo Finish
    Set grid = Range("Canvas")
    grid.ClearContents
    grid.Interior.ColorIndex = xlNone
    showStart = Clock()
    Dim act1 As Long, act2 As Long
    act1 = RunAct(1, "Act 1: one call to Excel per cell")
    If Not stopping Then act2 = RunAct(2, "Act 2: every cell in one call")
    Range("Countdown").Value2 = 0
    Range("Status").Value2 = IIf(stopping, "Stopped. ", "") & "Act 1 drew " & act1 & " frames, act 2 drew " & _
                             act2 & ". Same maths, same cells."
Finish:
    If Err.Number <> 0 Then Range("Status").Value2 = "Stopped by an error: " & Err.Description
    running = False
End Sub

' The Stop button.
Public Sub StopShow()
    stopping = True
End Sub

' Frames until the act's time is up; returns how many were drawn.
Private Function RunAct(ByVal act As Long, ByVal title As String) As Long
    actTitle = title
    actFrames = 0
    actStart = Clock()
    actEnd = actStart + ACT_SECONDS
    lastTick = 0
    If act = 2 Then grid.Interior.ColorIndex = xlNone
    Do While Clock() < actEnd And Not stopping
        ComputeField Clock() - showStart
        If act = 1 Then PaintOneCellAtATime Else PaintInOneCall
        actFrames = actFrames + 1
        Tick
    Loop
    RunAct = actFrames
End Function

' Every cell's colour for time t, 0 to 1.
Private Sub ComputeField(ByVal t As Double)
    Dim y As Long
    For y = 1 To GRID_ROWS
        PlasmaRow y, t
    Next
End Sub

' Four sine waves added together, one travelling in circles, then wrapped so the colours cycle.
Private Sub PlasmaRow(ByVal y As Long, ByVal t As Double)
    Dim x As Long, v As Double, cx As Double, cy As Double
    cy = y / GRID_ROWS - 0.5 + 0.35 * Cos(t / 3)
    For x = 1 To GRID_COLS
        cx = x / GRID_COLS - 0.5 + 0.35 * Sin(t / 2)
        v = Sin(x * 0.11 + t) + Sin((y * 0.13 + t) * 0.7) + Sin((x * 0.06 + y * 0.09 + t) * 0.9) _
          + Sin(Sqr(80 * (cx * cx + cy * cy) + 1) * 3 - t * 1.5)
        v = (v + 4) / 8 + t * 0.03
        v = v - Int(v)
        field(y, x) = Abs(2 * v - 1)            ' a triangle wave, so the cycle has no seam
    Next
End Sub

' Act 1: one call to Excel for every cell.
Private Sub PaintOneCellAtATime()
    Dim y As Long, x As Long
    For y = 1 To GRID_ROWS
        For x = 1 To GRID_COLS
            SetPixel y, x
        Next
        Tick
        If Clock() >= actEnd Or stopping Then Exit Sub
    Next
End Sub

Private Sub SetPixel(ByVal y As Long, ByVal x As Long)
    grid.Cells(y, x).Interior.Color = Palette(field(y, x))
End Sub

Private Function Palette(ByVal v As Double) As Long
    If v < 0.5 Then Palette = Mix(DEEP, PINK, v * 2) Else Palette = Mix(PINK, GOLD, v * 2 - 1)
End Function

' Part way from one colour to another, channel by channel.
Private Function Mix(ByVal a As Long, ByVal b As Long, ByVal f As Double) As Long
    Dim i As Long, part As Long, fromA As Long, toB As Long
    For i = 0 To 2
        part = 256 ^ i
        fromA = (a \ part) And &HFF
        toB = (b \ part) And &HFF
        Mix = Mix + CLng(fromA + (toB - fromA) * f) * part
    Next
End Function

' Act 2: every value in one call; the colour scale does the colouring.
Private Sub PaintInOneCall()
    grid.Value2 = field
End Sub

' The countdown and the frame rate, at most four times a second, then let Excel draw.
Private Sub Tick()
    Dim at As Double, secondsLeft As Long
    at = Clock()
    If at - lastTick >= 0.25 Then
        lastTick = at
        secondsLeft = -Int(-(showStart + 2 * ACT_SECONDS - at))
        If secondsLeft < 0 Then secondsLeft = 0
        If Range("Countdown").Value2 <> secondsLeft Then Range("Countdown").Value2 = secondsLeft
        Range("Status").Value2 = actTitle & ":  " & Format(actFrames / (at - actStart + 0.000001), "0.0") & _
                                 " frames a second"
    End If
    DoEvents
End Sub

' Seconds, from a clock that does not go back to 0 at midnight as Timer does.
Private Function Clock() As Double
    Clock = CDbl(Date) * 86400# + Timer
End Function
'@
    $ws = $d.Sheet
    $xl.ActiveWindow.Zoom = 100
    # Every cell 10 pixels square: 7.5 points high, and as many characters wide as measures 7.5 points.
    $ws.Cells.RowHeight = 7.5
    $cw = 0.3
    $ws.Columns(1).ColumnWidth = $cw
    while ($ws.Columns(1).Width -lt 7.49 -and $cw -lt 3) { $cw += 0.01; $ws.Columns(1).ColumnWidth = $cw }
    $ws.Cells.ColumnWidth = $cw

    # The plasma: 96 by 54 cells, values hidden, coloured by a three-colour scale from 0 to 1.
    $grid = $ws.Range('B14').Resize(54, 96)
    $grid.Name = 'Canvas'
    $grid.NumberFormat = ';;;'
    $scale = $grid.FormatConditions.AddColorScale(3)
    $stops = @(@(0, 0x501014), @(0.5, 0x8C32D6), @(1, 0x5AD6FF))
    for ($i = 0; $i -lt 3; $i++) {
        $c = $scale.ColorScaleCriteria.Item($i + 1)
        $c.Type = 0                                 # xlConditionValueNumber
        # a Variant property: set through InvokeMember, as Set-Table does, since the binder refuses a number
        [void][System.__ComObject].InvokeMember('Value', [Reflection.BindingFlags]::SetProperty, $null, $c,
                                                [object[]]@([double]$stops[$i][0]))
        $c.FormatColor.Color = $stops[$i][1]
        Clear-ComRef @($c)
    }

    # Above it: the countdown, a status line, and the two buttons.
    # named by their top-left cells: a merged range's Value2 is an array
    $cd = $ws.Range('B2:K11'); $cd.Merge(); $ws.Range('B2').Name = 'Countdown'; $cd.Value2 = 20
    $cd.Font.Size = 40; $cd.Font.Bold = $true; $cd.HorizontalAlignment = -4108; $cd.VerticalAlignment = -4108
    $st = $ws.Range('M2:CS6'); $st.Merge(); $ws.Range('M2').Name = 'Status'
    $st.Value2 = 'Press Start the show. It runs for 20 seconds.'
    $st.Font.Size = 14; $st.Font.Bold = $true; $st.VerticalAlignment = -4108
    Add-Button $ws 'StartTheShow' 'Start the show' 'StartShow' 'M8'
    Add-Button $ws 'StopTheShow' 'Stop' 'StopShow' 'AC8'

    # Beside it: what to do, and what to look for in Perfetto.
    $guide = $ws.Shapes.AddTextbox(1, $ws.Range('CU2').Left, $ws.Range('CU2').Top, 380, 650)
    $guide.Line.Visible = 0
    $text = $guide.TextFrame2.TextRange
    $text.Text = (@(
        '10  A plasma, painted two ways',
        'The same VBA maths drawn twice: first one call to Excel per cell, then every cell in one call. Perfetto shows where the time goes.',
        '',
        'Try it',
        '1.  Developer tab > XRayXL > Arm.',
        '2.  Press Start the show, and watch for 20 seconds. Stop ends it early.',
        '3.  Developer tab > XRayXL > Disarm.',
        '4.  Press Perfetto, then Open Trace in Perfetto, and answer Yes.',
        '',
        'Look for this in Perfetto',
        '-  StartShow, holding one RunAct for each act.',
        '-  Act 1: each frame is a wide PaintOneCellAtATime, a comb of SetPixel calls, each calling Palette and then Mix. ComputeField beside it is a sliver.',
        '-  Act 2: the frames are thin: ComputeField with its 54 PlasmaRow calls, then one PaintInOneCall.',
        '-  Excel''s SheetChange events, marking each write.',
        '-  Click a SetPixel: most of its time is its own, spent in Excel colouring one cell.'
    ) -join "`r")
    $text.Font.Size = 10
    $text.Font.Name = 'Segoe UI'
    $text.Paragraphs(1).Font.Size = 16; $text.Paragraphs(1).Font.Bold = -1
    $text.Paragraphs(2).Font.Italic = -1
    $text.Paragraphs(4).Font.Bold = -1
    $text.Paragraphs(10).Font.Bold = -1
    Clear-ComRef @($text, $guide, $st, $cd, $scale, $grid)
    Save-As $d '10_Plasma.xlsm'
}
finally {
    # dropping the variable releases nothing: Excel is a ref-counted COM server
    if ($xl) {
        try { $xl.Quit() } catch {}
        Clear-ComRef @($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # a failed build can leave references Quit cannot outlive; end our own Excel, by its pid
    Start-Sleep -Seconds 2
    if ($xlPid -and (Get-Process -Id $xlPid -ErrorAction SilentlyContinue)) { Stop-Process -Id $xlPid -Force }
}

foreach ($leaf in $retired) {
    $old = Join-Path $wbOut $leaf
    if (Test-Path $old) { Remove-Item $old; Write-Host "  retired -> $leaf" }
}
Write-Host "workbooks -> $wbOut"
Write-Host "done."
