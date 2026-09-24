#pragma once
#include <windows.h>
#include <cstring>

namespace core
{
    // A per-thread frame stack that grows on the heap, so no call depth goes unrecorded. No
    // constructor, so it can live in __declspec(thread) storage zero-filled by the loader.
    // T must be trivially copyable: growing moves it with memcpy.
    template <class T, int kFirst>
    struct TlsStack
    {
        T*  items;
        int capacity;

        // Room for index `i`. False only when the heap refuses. Growing moves the items, so no
        // reference into the stack may be held across a call that can push.
        bool Reserve(int i)
        {
            if (i < capacity) return true;
            int cap = capacity ? capacity : kFirst;
            while (cap <= i) cap *= 2;
            T* grown = static_cast<T*>(HeapAlloc(GetProcessHeap(), 0, sizeof(T) * cap));
            if (!grown) return false;
            if (items)
            {
                memcpy(grown, items, sizeof(T) * capacity);
                HeapFree(GetProcessHeap(), 0, items);
            }
            items = grown;
            capacity = cap;
            return true;
        }

        // At thread exit only, when no hook can be running on the thread.
        void Release()
        {
            if (items) HeapFree(GetProcessHeap(), 0, items);
            items = nullptr;
            capacity = 0;
        }

        T&       operator[](int i)       { return items[i]; }
        const T& operator[](int i) const { return items[i]; }
    };
}
