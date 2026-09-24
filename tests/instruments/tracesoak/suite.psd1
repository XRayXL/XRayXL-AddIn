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

    # Warm-up, two measured phases and three workbooks to build; XRAY_SOAK_SECONDS
    # lengthens the phases, so leave generous headroom.
    TestTimeoutSeconds = 3000
    ModulesOfInterest = @('xrayxl', 'vbe7')

    # Otherwise the add-in writes to the shared %TEMP%\XRayXL, keyed by recycled pids, where
    # files pile up. tests\sweep\suite.psd1 sets it only for tests\sweep\.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'
    }

    # Same trap, same guard -- see tests\sweep\suite.psd1.
    RequireNotOlderThan = @{
        '..\..\..\build\addin\XRayXL64.xll' = '..\..\..\build\x64\Release\XRayXL\XRayXL64.xll'
        '..\..\..\build\x64\Release\XRayXL\XRayXL64.xll' = '..\..\..\src'
    }
}
