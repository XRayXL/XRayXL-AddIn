#include "rowjson.h"
#include "core/jsonvalue.h"

namespace emit
{
namespace json
{
    namespace
    {
        enum class Kind { Text, Integer, Json };
        struct Field { const char* key; const char* value; Kind kind; };

        bool IsInteger(const char* s)
        {
            if (*s == '-') ++s;
            if (!*s) return false;
            for (; *s; ++s) if (*s < '0' || *s > '9') return false;
            return true;
        }
    }

    void AppendRow(const csv::Row& r, core::TextBuf& out)
    {
        const Field fields[] = {
            { "kind",      r.kind,      Kind::Text    }, { "source",   r.source,   Kind::Text    },
            { "span",      r.span,      Kind::Integer }, { "parent",   r.parent,   Kind::Integer },
            { "depth",     r.depth,     Kind::Integer }, { "thread",   r.thread,   Kind::Integer },
            { "qpc",       r.qpc,       Kind::Integer }, { "module",   r.module,   Kind::Text    },
            { "function",  r.function,  Kind::Text    }, { "proc",     r.proc,     Kind::Text    },
            { "typetext",  r.typetext,  Kind::Text    }, { "caller",   r.caller,   Kind::Text    },
            { "callerref", r.callerref, Kind::Text    }, { "argcount", r.argcount, Kind::Integer },
            { "args",      r.args,      Kind::Json    }, { "ret",      r.ret,      Kind::Json    },
            { "rettype",   r.rettype,   Kind::Text    }, { "outcome",  r.outcome,  Kind::Text    },
            { "ticks",     r.ticks,     Kind::Integer }, { "trust",    r.trust,    Kind::Text    },
        };
        bool first = true;
        for (const Field& f : fields)
        {
            if (!f.value || !f.value[0]) continue;
            if (!first) out.Append(',');
            first = false;
            core::AppendJsonString(out, f.key);
            out.Append(':');
            if (f.kind == Kind::Json || (f.kind == Kind::Integer && IsInteger(f.value)))
                out.Append(f.value);
            else
                core::AppendJsonString(out, f.value);
        }
        out.Append("}\r\n");
    }
}
}   // namespace emit
