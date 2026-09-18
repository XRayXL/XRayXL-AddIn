@{
    # Its own folder because it needs its own session: one is shared only between tests with the same settings.
    RegisterXll = @(
        '..\..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 180
    SessionEnvironment = @{ XRAYXL_RIBBON = '0' }
}
