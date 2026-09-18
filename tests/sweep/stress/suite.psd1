@{
    RegisterXll = @('..\..\..\build\addin\XRayXL64.xll')
    RegisterXllSettleSeconds = 5
    # The slowest stress case has a 60s budget of its own; 600 covers the worst case without
    # loosening anyone else's run.
    TestTimeoutSeconds = 600
}
