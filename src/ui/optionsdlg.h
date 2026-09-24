#pragma once

// The Options dialog, drawn to match Excel's own. Settings go through ui::ribbon::model.

namespace ui
{
namespace options
{
    // Modal against `ownerHwnd` (null: unowned). True if OK applied the settings.
    bool Show(void* ownerHwnd);
}
}
