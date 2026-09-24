@{
    # The traced add-in first, so its UDFs resolve in saved formulas. Registering
    # does not arm: each test arms explicitly.
    RegisterXll = @(
        '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 300
}
