#pragma once

// The Diagnostics dialog: what is loaded into this Excel, what the environment says, and what
// the process is using. Read-only, so it opens armed or not.

namespace ui
{
namespace diagnostics
{
    // Modal against `owner` (null: Excel's main window). False if it could not be opened.
    bool Show(void* ownerHwnd);
}
}
