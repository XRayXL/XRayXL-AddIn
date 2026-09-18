#pragma once
#include "xll/xlltrace.h"      // ArmReport

// THE SESSION: arming and disarming BOTH sources in the one order that is
// safe, and the exports Excel calls to do it. This is the layer above xll/
// and vba/, which do not know about each other -- everything they share is
// in core/, and everything that drives both is here.
namespace app
{
    // "0.1.0", from version.props via the resource defines.
    const char* VersionText();

    // VBA first and independently, then the XLL side: the two are separate
    // sources, and a workbook with VBA and no XLL add-ins must still trace.
    xll::ArmReport Arm();

    // The output ring's loss report, the VBA side unconditionally (it can be
    // armed when the XLL side is not), then the XLL side. Idempotent: it
    // arrives from Application.Run and from xlAutoClose, either first.
    // shuttingDown skips the coverage check, which needs the object model.
    void Disarm(bool shuttingDown = false);

    // Either source armed -- what the setters test and XRayXL_IsArmed answers.
    bool IsArmed();

    // The XRayXL_* commands and functions, registered with Excel at load.
    void RegisterCommands();
}
