@{
    # For hangwhere.py --ours triage of any dump these suites produce: the
    # modules whose presence on a parked thread's stack is the question.
    ModulesOfInterest = @('xrayxl', 'vbe7')

    # THE DEPLOYED ADD-IN MUST NOT BE OLDER THAN THE BUILD IT COMES FROM.
    # Every suite here registers ..\..\..\build\addin\XRayXL64.xll, which only
    # tools\deploy.ps1 writes. Skipping that step fails nothing -- the run
    # simply certifies the previous binary, in green -- so the run refuses to start.
    RequireNotOlderThan = @{
        '..\..\build\addin\XRayXL64.xll' = '..\..\build\x64\Release\XRayXL\XRayXL64.xll'
        # ...and the BUILD must not be older than the SOURCE it came from, or a
        # build that FAILED leaves the previous binary for the deploy to copy,
        # the two timestamps agree, and the run certifies code never compiled.
        '..\..\build\x64\Release\XRayXL\XRayXL64.xll' = '..\..\src'
        # The add-in the XLL suites trace: built from its own folder and the shared XLL header.
        '..\..\build\x64\Release\TracedAddin\TracedAddin64.xll' = @('..\fixtures\TracedAddin', '..\fixtures\xll_common')
    }

    # EVERY SESSION WRITES INTO ITS OWN DIRECTORY, beside the run. The XLL
    # takes its output root from XRAYXL_OUTPUT_DIR (logs, trace files, the
    # DIAG corpus all go under it); without this they pile up in %TEMP%\XRayXL
    # keyed by pid, and a reused pid reads an old log. The suites' helpers (Get-XRayPaths, Get-XRayTraceCsv)
    # follow the same variable.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'
    }
}
