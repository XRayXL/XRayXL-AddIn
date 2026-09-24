@{
    # For hangwhere.py --ours triage of any dump these suites produce: the
    # modules whose presence on a parked thread's stack is the question.
    ModulesOfInterest = @('xrayxl', 'vbe7')

    # Only tools\deploy.ps1 writes build\addin\XRayXL64.xll; skipping it would certify the
    # previous binary in green, so a stale copy stops the run.
    RequireNotOlderThan = @{
        '..\..\build\addin\XRayXL64.xll' = '..\..\build\x64\Release\XRayXL\XRayXL64.xll'
        # A failed build leaves the previous binary for the deploy to copy.
        '..\..\build\x64\Release\XRayXL\XRayXL64.xll' = '..\..\src'
        # The add-in the XLL suites trace: built from its own folder and the shared XLL header.
        '..\..\build\x64\Release\TracedAddin\TracedAddin64.xll' = @('..\fixtures\TracedAddin', '..\fixtures\xll_common')
    }

    # Each session writes its logs and traces into its own directory: in the shared %TEMP%\XRayXL,
    # keyed by pid, a reused pid would read an old log. The suites' helpers follow the same variable.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'
    }
}
