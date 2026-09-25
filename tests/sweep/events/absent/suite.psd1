@{
    # Its own folder for its own session: the diagnostic switch makes this Excel pretend it lacks two events.
    RegisterXll = @('..\..\..\..\build\addin\XRayXL64.xll')
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 180
    SessionEnvironment = @{ XRAYXL_DIAG = '1'; XRAYXL_EVENTS_ABSENT = 'SheetTableUpdate,WorkbookModelChange' }
}
