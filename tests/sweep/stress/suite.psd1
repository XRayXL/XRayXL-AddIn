@{
    RegisterXll = @('..\..\..\build\addin\XRayXL64.xll')
    RegisterXllSettleSeconds = 5
    # The slowest stress case has a 60s budget of its own; x10 was the old
    # cap. 600 covers the worst case without loosening anyone else's run.
    TestTimeoutSeconds = 600
}
