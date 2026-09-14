@{
    # The traced add-in first (its UDFs must resolve in saved formulas), the
    # product second. Registration happens once per session, by the worker,
    # before the test runs; arming stays the test's own explicit act.
    RegisterXll = @(
        '..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 300
}
