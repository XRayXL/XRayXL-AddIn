. (Join-Path $PSScriptRoot '..\_strings.ps1')
# FALSE is easy to lose silently: an empty field and a missing argument look alike
$case = @{ F = "=TxA(FALSE)";         W = 'boolean false'
           Value = '0'; Args = @('a1:A=0') }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-FuzzCase $case
