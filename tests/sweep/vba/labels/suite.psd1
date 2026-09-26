@{
    # Its own folder for its own session: diagnostics on, so each signature names the opcode that
    # typed it and the test can tell a label from a load. Keeps the root's per-session output.
    SessionEnvironment = @{
        XRAYXL_OUTPUT_DIR = '{SessionDir}'
        XRAYXL_DIAG = '1'
    }
}
