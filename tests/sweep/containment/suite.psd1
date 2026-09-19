@{
    # Its own folder because it needs its own session: XRayXL_FaultProbe is registered only with
    # diagnostics on, and a session is shared only between tests with the same settings.
    RegisterXll = @(
        '..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 180
    SessionEnvironment = @{ XRAYXL_DIAG = '1' }
}
