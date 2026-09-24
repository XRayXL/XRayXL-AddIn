@{
    # Console programs built from the product's own sources: each runs with no Excel,
    # so the session the manager starts around it goes unused.
    TestTimeoutSeconds = 300

    # An exe older than the source it is built from would certify code that was
    # never compiled, so the run refuses to start. These tests load no add-in,
    # so this replaces the root's add-in check rather than adding to it.
    RequireNotOlderThan = @{
        '..\..\..\build\x64\Release\unit\callerdecode.exe' = @('..\..\..\src', 'callerdecode')
        '..\..\..\build\x64\Release\unit\csv_close_race_test.exe' = @('..\..\..\src', 'csv_close_race_test')
        '..\..\..\build\x64\Release\unit\diagnosticsdlg_test.exe' = @('..\..\..\src', 'diagnosticsdlg_test')
        '..\..\..\build\x64\Release\unit\optionsdlg_test.exe' = @('..\..\..\src', 'optionsdlg_test')
        '..\..\..\build\x64\Release\unit\decode_test.exe' = @('..\..\..\src', 'decode_test')
        '..\..\..\build\x64\Release\unit\paramparse_test.exe' = @('..\..\..\src', 'paramparse_test')
        '..\..\..\build\x64\Release\unit\pcode_scan_guard_test.exe' = @('..\..\..\src', 'pcode_scan_guard_test')
        '..\..\..\build\x64\Release\unit\proctable_test.exe' = @('..\..\..\src', 'proctable_test')
        '..\..\..\build\x64\Release\unit\ribbonmodel_test.exe' = @('..\..\..\src', 'ribbonmodel_test')
        '..\..\..\build\x64\Release\unit\ring_stress.exe' = @('..\..\..\src', 'ring_stress')
        '..\..\..\build\x64\Release\unit\rowcsv_test.exe' = @('..\..\..\src', 'rowcsv_test')
        '..\..\..\build\x64\Release\unit\vbaderive_roles_test.exe' = @('..\..\..\src', 'vbaderive_roles_test')
        '..\..\..\build\x64\Release\unit\xlltypeplan_test.exe' = @('..\..\..\src', 'xlltypeplan_test')
    }
}