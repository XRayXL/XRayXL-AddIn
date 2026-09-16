@{
    # THE RING SUITE. Cases that only the buffered output path can
    # fail: one proves a healthy buffer loses nothing and keeps file order; one
    # starves a deliberately tiny ring under BUFFERWHENFULL=DROP and proves the loss is
    # marked once and reconciles; one starves the same ring under BUFFERWHENFULL=PAUSE
    # and proves it loses nothing at the cost of pausing the hot path. The rest
    # of the tree runs the shipped default and would never notice a
    # drop-accounting bug -- this is where it must show.
    RegisterXll = @('..\..\..\build\addin\XRayXL64.xll')
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 300
}
