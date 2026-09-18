#pragma once

// The Options dialog, drawn to match Excel's own. Settings go through ui::ribbon::model.

namespace ui
{
namespace options
{
    // Modal against `owner` (null: Excel's main window). True if OK applied the settings.
    bool Show(void* ownerHwnd);
}
}
