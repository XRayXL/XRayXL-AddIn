# A 255-character name (VBA's identifier limit), asserted whole: a non-blank
# check would pass on a truncated name.
$case = @{ Name='very-long-procedure-names'
     Modules=@{
       'M'=("Public Sub {0}()`r`n    Dim z As Long`r`n    z = 1`r`nEnd Sub`r`nPublic Sub Go()`r`n    {0}`r`nEnd Sub`r`n" -f ('A' + ('b' * 254)))
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $want = 'A' + ('b' * 254)
        $names = @(Get-VbaEntryNames $t)
        $long = @($names | Where-Object { $_ -like 'Ab*' })
        if ($long.Count -lt 1) { return "the long-named procedure was not named at all; saw [$($names -join ',')]" }
        $got = $long[0]
        if ($got -like '*~') {
            return "name truncated: $($got.Length) of $($want.Length) characters, ending '~'" }
        if ($got -ne $want) {
            return "name is $($got.Length) characters, expected $($want.Length)" }
        $null }
     Why='VBA allows a 255-character identifier and the name buffers are sized to
          match; asserts the WHOLE name survives, because the previous assertion
          (merely not blank) passed on a name cut at 95' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
