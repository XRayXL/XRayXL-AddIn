#include "notify.h"
#include <windows.h>

namespace core
{
    namespace
    {
        // installed once at load, but fired from whatever thread Application.Run arrived on
        volatile LONG64 g_fn = 0;
    }

    void SubscribeStateChanged(StateChangedFn fn)
    {
        InterlockedExchange64(&g_fn, reinterpret_cast<LONG64>(fn));
    }

    void NotifyStateChanged()
    {
        const LONG64 fn = InterlockedCompareExchange64(&g_fn, 0, 0);
        if (fn) reinterpret_cast<StateChangedFn>(fn)();
    }
}
