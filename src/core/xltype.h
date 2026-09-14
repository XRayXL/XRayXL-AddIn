#pragma once
#include "xlcall.h"

namespace core
{
    // xltype without the free flags.
    constexpr int kXlTypeMask = ~(xlbitXLFree | xlbitDLLFree);

    // Excel12 function number without its flags.
    constexpr int kXlFunctionMask = ~(xlCommand | xlSpecial | xlIntl | xlPrompt);
}
