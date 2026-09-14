@{
    # An INSTRUMENT, deliberately outside suites\ so a full sweep never runs
    # it: it measures a TREND rather than defending a behaviour, it takes
    # minutes, and it has no pass threshold to assert -- the numbers are the
    # result. Same reasoning as tests\vbahammer.
    #   .\StretchXL\StretchXL.ps1 -Parallel 1 -Path .\tests\tracesoak -OutDir <dir>
    #
    # BOTH sources, because the question is what ONE arming costs over a long
    # session and the answer must include the XLL hooks and the VBA dispatch
    # patches together. The traced add-in first: its UDFs have to resolve in
    # the saved formulas.
    RegisterXll = @(
        '..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\build\addin\XRayXL64.xll'
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
    # scans them all. suites\suite.psd1 sets it only for suites\.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'
    }

    # Same trap, same guard -- see suites\suite.psd1.
    RequireNotOlderThan = @{
        '..\..\build\addin\XRayXL64.xll' = '..\..\build\x64\Release\XRayXL\XRayXL64.xll'
        '..\..\build\x64\Release\XRayXL\XRayXL64.xll' = '..\..\src'
    }
}
