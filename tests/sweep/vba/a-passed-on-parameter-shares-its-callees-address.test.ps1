# A parameter the body only passes on carries its callee's address: it has no type of its own,
# so that is how a reader types it. The callee's VarPtr is the independent expected address.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public pTakeD As LongPtr, pTakeL As LongPtr, pTakeS As LongPtr, pTakeCopy As LongPtr
Private gD As Double, gL As Long, gS As String

Public Sub Drive()
    Dim d As Double, s As String
    d = 2.5: s = "abc"
    PassRef d
    PassVal 5
    PassStr s
    PassCopy d
End Sub
Public Function Ptrs() As String
    Ptrs = CStr(pTakeD) & ";" & CStr(pTakeL) & ";" & CStr(pTakeS) & ";" & CStr(pTakeCopy)
End Function

' Each passes x on and does nothing else with it.
Private Sub PassRef(x As Double)
    TakeD x
End Sub
Private Sub PassVal(ByVal n As Long)
    TakeL n
End Sub
Private Sub PassStr(x As String)
    TakeS x
End Sub
' Parentheses make a copy, so the callee's storage is not x.
Private Sub PassCopy(x As Double)
    TakeCopy (x)
End Sub

Private Sub TakeD(y As Double)
    gD = y: pTakeD = VarPtr(y)
End Sub
Private Sub TakeL(y As Long)
    gL = y: pTakeL = VarPtr(y)
End Sub
Private Sub TakeS(y As String)
    gS = y: pTakeS = VarPtr(y)
End Sub
Private Sub TakeCopy(y As Double)
    gD = y: pTakeCopy = VarPtr(y)
End Sub
'@

function A1Of($Rows, [string]$Fn) {
    $a = ArgsOf $Rows $Fn
    if ($a -match '^a1:([^=@]*)@0x([0-9A-F]+)=(.*)$') {
        return @{ type=$Matches[1]; addr=[Convert]::ToUInt64($Matches[2], 16); value=$Matches[3]; text=$a }
    }
    return @{ type=''; addr=$null; value=''; text=$a }
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    New-XRayMacroBook $sx 'PassedOn' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body {
        $app.Run($leaf + '!Drive') | Out-Null
        [string]$app.Run($leaf + '!Ptrs')
    }
    $vp = @(([string]$s.Result) -split ';' | ForEach-Object { [uint64][int64]$_ })

    $cases = @(
        @{ name='a-byref-parameter-passed-on';       parent='PassRef'; callee='TakeD'; ptr=$vp[0]; ptype='?none'; ctype='Double&'; cvalue='2.5' }
        @{ name='a-byval-parameter-passed-on';       parent='PassVal'; callee='TakeL'; ptr=$vp[1]; ptype='?none'; ctype='Long&';   cvalue='5' }
        @{ name='a-string-passed-on-keeps-its-value'; parent='PassStr'; callee='TakeS'; ptr=$vp[2]; ptype='?none'; ctype='String&'; cvalue='"abc"' }
    )
    foreach ($c in $cases) {
        $p = A1Of $s.Rows $c.parent
        $k = A1Of $s.Rows $c.callee
        Check $c.name ($p.type -eq $c.ptype -and $k.type -eq $c.ctype -and $k.value -eq $c.cvalue -and
                       $p.addr -eq $c.ptr -and $k.addr -eq $c.ptr) `
              ("$($c.parent) [$($p.text)]  $($c.callee) [$($k.text)]  VarPtr 0x{0:X}" -f $c.ptr)
    }
    $str = A1Of $s.Rows 'PassStr'
    Check 'the-passed-on-string-still-reads-as-its-value' ($str.value -eq '"abc"') "PassStr [$($str.text)]"

    $p = A1Of $s.Rows 'PassCopy'
    $k = A1Of $s.Rows 'TakeCopy'
    Check 'a-copy-in-parentheses-does-not-share-the-address' `
          ($p.type -eq 'Double&' -and $k.addr -eq $vp[3] -and $p.addr -ne $k.addr) `
          ("PassCopy [$($p.text)]  TakeCopy [$($k.text)]  VarPtr 0x{0:X}" -f $vp[3])

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("PassRef [{0}] TakeD [{1}]" -f (ArgsOf $s.Rows 'PassRef'), (ArgsOf $s.Rows 'TakeD'))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
