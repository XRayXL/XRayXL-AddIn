@{
    # An instrument, outside tests\sweep\ so a full sweep never runs it: it measures a trend,
    # takes minutes and has no pass threshold.
    #    .\StretchXL\StretchXL.ps1 -Parallel 1 -Path .\tests\instruments\tracesoak -OutDir <dir>
    #
    # Both sources, because one arming includes the XLL hooks and the VBA dispatch patches
    # together. The traced add-in first: its UDFs have to resolve in the saved formulas.
    RegisterXll = @(
        '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5

    # Warm-up plus two measured phases, plus building three workbooks. The
    # default shape is ~7 minutes; XRAY_SOAK_SECONDS raises the phases, so
    # leave generous headroom above it.
    TestTimeoutSeconds = 3000
    ModulesOfInterest = @('xrayxl', 'vbe7')

    # Every session writes into its own directory. Without this the add-in
    # writes to the shared %TEMP%\XRayXL, keyed by pids Windows recycles, where
    # files pile up without limit and every "newest trace for this pid" lookup
    # scans them all. tests\sweep\suite.psd1 sets it only for tests\sweep\.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'
    }

    # Same trap, same guard -- see tests\sweep\suite.psd1.
    RequireNotOlderThan = @{
        '..\..\..\build\addin\XRayXL64.xll' = '..\..\..\build\x64\Release\XRayXL\XRayXL64.xll'
        '..\..\..\build\x64\Release\XRayXL\XRayXL64.xll' = '..\..\..\src'
    }
}
