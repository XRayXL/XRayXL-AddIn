@{
    # An instrument, outside tests\sweep\ so a full sweep never runs it: it hunts a crash, takes
    # minutes and may kill its Excel.
    #    .\StretchXL\StretchXL.ps1 -Parallel 1 -Path .\tests\instruments\vbahammer -OutDir <dir>
    # Both sources, so a cycle arms the XLL hooks and the VBA dispatch patches together. Test
    # XLL first: its UDFs have to resolve in the saved formulas.
    RegisterXll = @(
        '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 3000
    ModulesOfInterest = @('xrayxl', 'vbe7')

    # Every session writes into its own directory. Without this the add-in
    # writes to the shared %TEMP%\XRayXL, keyed by pids Windows recycles, where
    # files pile up without limit and every "newest trace for this pid" lookup
    # scans them all. tests\sweep\suite.psd1 sets it only for tests\sweep\.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'

        # the dump switch is read once at add-in load, too early for the test to set
        XRAYXL_CRASHDUMP  = '1'
    }

    # Same trap, same guard -- see tests\sweep\suite.psd1.
    RequireNotOlderThan = @{
        '..\..\..\build\addin\XRayXL64.xll' = '..\..\..\build\x64\Release\XRayXL\XRayXL64.xll'
        # ...and the BUILD must not be older than the SOURCE it came from, or a
        # build that FAILED leaves the previous binary for the deploy to copy,
        # the two timestamps agree, and the run certifies code never compiled.
        '..\..\..\build\x64\Release\XRayXL\XRayXL64.xll' = '..\..\..\src'
    }
}
