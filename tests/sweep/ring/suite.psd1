@{
    # Cases only the buffered output path can fail. The rest of the tree runs the shipped
    # default and would never notice a drop-accounting bug.
    RegisterXll = @('..\..\..\build\addin\XRayXL64.xll')
    RegisterXllSettleSeconds = 5
    TestTimeoutSeconds = 300
}
