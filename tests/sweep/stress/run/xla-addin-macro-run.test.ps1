$case = @{ Name='xla-addin-macro-run'
     Modules=@{
       'M'=@'
Public Sub Go()
    Application.Run "AddinMac"
End Sub
'@
     }
     Deps=@(
       @{ Name='AddTwo'; IsAddin=$true; Modules=@{
         'MA'=@'
Public Sub AddinMac()
    AddinDeep 5
End Sub
Public Sub AddinDeep(ByVal n As Long)
    If n > 0 Then AddinDeep n - 1
End Sub
'@
       } }
     )
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'AddinMac','AddinDeep'; if ($why) { return $why }
        $null }
     # Go runs AddinMac, which recurses AddinDeep 5 down to 0: each call inside the one before.
     Calls=@(
        @{ Function='Go';       Depth='1'; Parent=-1; Outcome='returned' }
        @{ Function='AddinMac'; Depth='2'; Parent=0;  Outcome='returned' } ) +
           @(0..5 | ForEach-Object { @{ Function='AddinDeep'; Args="a1:Long=$(5 - $_)"; Depth="$($_ + 3)"; Parent=($_ + 1); Outcome='returned' } })
     Why='Application.Run against a macro living in an .xlam add-in' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
