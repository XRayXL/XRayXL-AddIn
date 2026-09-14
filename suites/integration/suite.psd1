@{
    # A cross-source test: VBA tracing AND one XLL UDF in the same recalc, so
    # the traced add-in (TxB) must resolve in the saved formula. That first, then
    # the product -- the order the xll suite uses for the same reason.
    RegisterXll = @(
        '..\..\build\x64\Release\TracedAddin\TracedAddin64.xll'
        '..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 300
}
