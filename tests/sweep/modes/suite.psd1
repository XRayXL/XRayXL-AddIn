@{
    # The traced add-in first (its UDFs must resolve in saved formulas), the
    # product second. These tests exercise the trace-mode surface, so
    # they need both.
    RegisterXll = @(
        '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 300
}
