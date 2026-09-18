// A one-slot signal that armed state or a trace parameter changed.
#pragma once

namespace core
{
    using StateChangedFn = void (*)();

    // One subscriber; nullptr unsubscribes.
    void SubscribeStateChanged(StateChangedFn fn);

    // Runs the subscriber on the calling thread.
    void NotifyStateChanged();
}
