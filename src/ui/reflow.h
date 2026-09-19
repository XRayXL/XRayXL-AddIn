#pragma once
#include <string>

namespace ui
{
namespace text
{
    // A plain-text file's hard-wrapped lines rejoined, so a box that wraps can flow them. A blank
    // line ends a paragraph. A rule line of = or - and a numbered clause start their own line, and
    // a rule is drawn short so it never wraps; the line after a rule, a line indented as a block,
    // or one ending in an address also starts its own line. Line ends come out CRLF.
    inline std::string Reflow(const std::string& all)
    {
        std::string out;
        bool inParagraph = false, breakAfter = false;
        for (size_t at = 0; at < all.size(); )
        {
            size_t end = all.find('\n', at);
            if (end == std::string::npos) end = all.size();
            const std::string raw = all.substr(at, end - at);
            at = end + 1;

            const size_t first = raw.find_first_not_of(" \t\r"), last = raw.find_last_not_of(" \t\r");
            if (first == std::string::npos)
            {
                if (inParagraph) out += "\r\n\r\n";
                inParagraph = false;
                continue;
            }
            const std::string text = raw.substr(first, last - first + 1);
            const bool rule = text.size() >= 3 &&
                              (text[0] == '=' || text[0] == '-') &&
                              text.find_first_not_of(text[0]) == std::string::npos;
            const size_t digits = text.find_first_not_of("0123456789");
            const bool clause = digits > 0 && digits != std::string::npos && text[digits] == '.';
            if (inParagraph) out += (breakAfter || rule || clause) ? "\r\n" : " ";
            out += rule ? std::string(24, text[0]) : text;
            inParagraph = true;
            breakAfter = first >= 10 || raw[last] == '>' || rule;
        }
        return out;
    }
}
}
