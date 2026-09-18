@{
    # The product only: these tests are about the ribbon add-in's lifecycle, so nothing needs tracing.
    RegisterXll = @(
        '..\..\..\build\addin\XRayXL64.xll'
    )
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 180
}
