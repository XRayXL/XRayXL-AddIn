# Runs one unit test that XRayXL.sln builds, with no Excel, and reports each of its
# checks as a case. The exe's exit code and its failed checks decide the verdict.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')

function Invoke-UnitTest {
    param(
        [Parameter(Mandatory)][string]$Name,
        # Below the suite's TestTimeoutSeconds, so a hang is reported here rather than killed from outside.
        [int]$TimeoutSeconds = 240
    )
    try {
        $exe = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\..\..\build\x64\Release\unit\$Name.exe"))
        if (-not (Test-Path -LiteralPath $exe)) { Complete-Test -Fail -Detail "not built: $exe -- build XRayXL.sln" }

        $work = if ($env:STRETCH_SESSION_DIR) { $env:STRETCH_SESSION_DIR } else { [IO.Path]::GetTempPath() }
        $psi = New-Object Diagnostics.ProcessStartInfo $exe
        $psi.WorkingDirectory = $work
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true

        $sw = [Diagnostics.Stopwatch]::StartNew()
        $p = [Diagnostics.Process]::Start($psi)
        $stdout = $p.StandardOutput.ReadToEndAsync()
        $stderr = $p.StandardError.ReadToEndAsync()
        if (-not $p.WaitForExit($TimeoutSeconds * 1000)) {
            $p.Kill()   # our own child, by handle
            Complete-Test -Fail -Detail "$Name did not finish within $TimeoutSeconds s"
        }
        $p.WaitForExit()

        $passed = 0
        $failed = @()
        foreach ($line in ($stdout.Result -split "`r?`n" | Where-Object { $_ -ne '' })) {
            Write-Output "  $line"
            # [PASS]/[FAIL] lines, or callerdecode's "ok"/"FAIL" columns.
            $m = [regex]::Match($line, '^\s*(?:\[(PASS|FAIL)\]|(ok|FAIL)\s)\s*(.+?)\s*$')
            if (-not $m.Success) { continue }
            $text = $m.Groups[3].Value
            # The label before any padded column, so a case keeps its name when counts vary.
            $case = (($text -split '\s{2,}')[0].ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
            if ($case.Length -gt 60) { $case = $case.Substring(0, 60).TrimEnd('-') }
            if ($m.Groups[1].Value -eq 'PASS' -or $m.Groups[2].Value -eq 'ok') {
                $passed++
                Write-TestCase -Name $case -Pass
            }
            else {
                $failed += $text
                Write-TestCase -Name $case -Fail -Detail $text
            }
        }
        if ($stderr.Result.Trim()) { Write-Output "  stderr: $($stderr.Result.Trim())" }

        $secs = '{0:n1}' -f $sw.Elapsed.TotalSeconds
        if ($p.ExitCode -eq 0 -and $failed.Count -eq 0 -and $passed -gt 0) {
            Complete-Test -Pass -Detail "$passed checks passed in $secs s"
        }
        elseif ($failed.Count -gt 0) {
            Complete-Test -Fail -Detail "$($failed.Count) of $($passed + $failed.Count) checks failed (exit $($p.ExitCode)); first: $($failed[0])"
        }
        elseif ($p.ExitCode -ne 0) {
            Complete-Test -Fail -Detail "$Name exited $($p.ExitCode) without reporting a failed check"
        }
        else {
            Complete-Test -Fail -Detail "$Name reported no checks"
        }
    }
    catch {
        Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
    }
}
