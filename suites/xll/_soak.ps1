# The thread-safe XLL soak, shared by one test per argument type: a function in 20,000
# cells, recalculated 20 times on 8 threads. Every cell's value and every trace row is
# checked against the cell's own address, so a row attributed to the wrong cell, thread
# or call cannot pass.

# Streamed: 800,000 rows are far too many for Import-Csv.
if (-not ('XRaySoakCheck' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public class XRaySoakResult
{
    public List<string> Problems = new List<string>();
    public long ProblemCount;
    public long Lines;
    public long Entries;
    public long Exits;
    public int Threads;
    public void Add(string problem)
    {
        ProblemCount++;
        if (Problems.Count < 8) Problems.Add(problem);
    }
}

public static class XRaySoakCheck
{
    // One row is one line: the writer never puts CR or LF inside a field.
    static List<string> Fields(string line)
    {
        List<string> f = new List<string>(24);
        StringBuilder sb = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++)
        {
            char ch = line[i];
            if (quoted)
            {
                if (ch != '"') sb.Append(ch);
                else if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                else quoted = false;
            }
            else if (ch == '"') quoted = true;
            else if (ch == ',') { f.Add(sb.ToString()); sb.Length = 0; }
            else sb.Append(ch);
        }
        f.Add(sb.ToString());
        return f;
    }

    // "[Book.xlsx]S1!C17" -> row 17, column 3.
    static bool CellOf(string callerref, out int row, out int col)
    {
        row = 0; col = 0;
        int bang = callerref.LastIndexOf('!');
        if (bang < 0) return false;
        string a = callerref.Substring(bang + 1);
        int i = 0;
        while (i < a.Length && a[i] >= 'A' && a[i] <= 'Z') { col = col * 26 + (a[i] - 'A' + 1); i++; }
        return i > 0 && int.TryParse(a.Substring(i), out row) && row > 0;
    }

    // What a cell shows for =Fn(ROW()-offset,COLUMN()-offset). With natural set the add-in rejects
    // anything below 1 and names the argument, spelt exactly as tracedaddin.cpp spells it.
    public static string Expected(int row, int col, int offset, bool natural)
    {
        int r = row - offset, c = col - offset;
        if (!natural || (r >= 1 && c >= 1)) return r + ":" + c;
        string rw = r >= 1 ? null : "row " + r + " is not a natural number";
        string cw = c >= 1 ? null : "column " + c + " is not a natural number";
        if (rw != null && cw != null) return rw + "; " + cw;
        return rw ?? cw;
    }

    // argFormat takes {0} and {1}, the row and column arguments, e.g. "a1:B={0} a2:B={1}".
    public static XRaySoakResult Run(string path, string header, string book, string fn,
                                     int rows, int cols, int passes, string argFormat, string typeText,
                                     string retType, int offset, bool natural)
    {
        XRaySoakResult res = new XRaySoakResult();
        int[,] entries = new int[rows + 1, cols + 1];
        int[,] exits = new int[rows + 1, cols + 1];
        Dictionary<string, Tuple<int, int, string>> open = new Dictionary<string, Tuple<int, int, string>>();
        HashSet<string> threads = new HashSet<string>();
        List<long> inputs = new List<long>();
        string tag = "[" + book + "]";
        long lastSeq = 0;

        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
        {
            string first = sr.ReadLine();
            if (first != header) { res.Add("header is not the expected one: " + first); return res; }
            string[] names = header.Split(',');
            Dictionary<string, int> ix = new Dictionary<string, int>();
            for (int i = 0; i < names.Length; i++) ix[names[i]] = i;

            string line;
            while ((line = sr.ReadLine()) != null)
            {
                res.Lines++;
                List<string> f = Fields(line);
                if (f.Count != names.Length) { res.Add("line " + (res.Lines + 1) + " has " + f.Count + " fields"); continue; }

                long seq, input;
                if (!long.TryParse(f[ix["seq"]], out seq) || seq != lastSeq + 1)
                    res.Add("seq " + f[ix["seq"]] + " follows " + lastSeq);
                lastSeq = seq;
                if (long.TryParse(f[ix["input"]], out input)) inputs.Add(input);
                else res.Add("seq " + seq + ": input '" + f[ix["input"]] + "'");

                if (f[ix["source"]] != "XLL" || f[ix["function"]] != fn) continue;
                string kind = f[ix["kind"]], span = f[ix["span"]], thread = f[ix["thread"]];

                if (kind == "entry")
                {
                    int r, c;
                    string where = f[ix["callerref"]];
                    if (where.IndexOf(tag, StringComparison.Ordinal) < 0)
                    { res.Add("seq " + seq + ": entry from another book: " + where); continue; }
                    if (f[ix["caller"]] != "cell" || !CellOf(where, out r, out c) || r > rows || c > cols)
                    { res.Add("seq " + seq + ": caller " + f[ix["caller"]] + " '" + where + "' is not a cell of the grid"); continue; }
                    res.Entries++;
                    entries[r, c]++;
                    threads.Add(thread);
                    string args = string.Format(argFormat, r - offset, c - offset);
                    if (f[ix["args"]] != args) res.Add("seq " + seq + ": " + where + " args '" + f[ix["args"]] + "', expected '" + args + "'");
                    if (f[ix["argcount"]] != "2") res.Add("seq " + seq + ": argcount '" + f[ix["argcount"]] + "'");
                    if (f[ix["typetext"]] != typeText) res.Add("seq " + seq + ": typetext '" + f[ix["typetext"]] + "', expected '" + typeText + "'");
                    if (open.ContainsKey(span)) res.Add("seq " + seq + ": span " + span + " opened twice");
                    open[span] = Tuple.Create(r, c, thread);
                }
                else if (kind == "exit")
                {
                    Tuple<int, int, string> o;
                    if (!open.TryGetValue(span, out o)) { res.Add("seq " + seq + ": exit span " + span + " has no open entry"); continue; }
                    open.Remove(span);
                    res.Exits++;
                    exits[o.Item1, o.Item2]++;
                    if (thread != o.Item3) res.Add("seq " + seq + ": span " + span + " entered on thread " + o.Item3 + ", exited on " + thread);
                    string ret = "\"" + Expected(o.Item1, o.Item2, offset, natural) + "\"";
                    if (f[ix["ret"]] != ret) res.Add("seq " + seq + ": ret '" + f[ix["ret"]] + "', expected '" + ret + "'");
                    if (f[ix["rettype"]] != retType) res.Add("seq " + seq + ": rettype '" + f[ix["rettype"]] + "', expected '" + retType + "'");
                    if (f[ix["outcome"]] != "returned") res.Add("seq " + seq + ": outcome '" + f[ix["outcome"]] + "'");
                    if (f[ix["trust"]] != "exit") res.Add("seq " + seq + ": trust '" + f[ix["trust"]] + "'");
                    long ticks;
                    if (!long.TryParse(f[ix["ticks"]], out ticks) || ticks < 0) res.Add("seq " + seq + ": ticks '" + f[ix["ticks"]] + "'");
                }
            }
        }

        if (open.Count > 0) res.Add(open.Count + " entr(ies) never exited");
        int badCells = 0; string firstBad = null;
        for (int r = 1; r <= rows; r++)
            for (int c = 1; c <= cols; c++)
                if (entries[r, c] != passes || exits[r, c] != passes)
                {
                    badCells++;
                    if (firstBad == null) firstBad = "R" + r + "C" + c + " entries=" + entries[r, c] + " exits=" + exits[r, c];
                }
        if (badCells > 0) res.Add(badCells + " cell(s) not called exactly " + passes + " times, e.g. " + firstBad);

        inputs.Sort();
        for (int i = 1; i < inputs.Count; i++)
            if (inputs[i] == inputs[i - 1]) { res.Add("input " + inputs[i] + " used twice"); break; }
        if (inputs.Count > 0)
        {
            long holes = (inputs[inputs.Count - 1] - inputs[0] + 1) - inputs.Count;
            if (holes != 0) res.Add(holes + " input number(s) missing: rows were lost");
        }
        res.Threads = threads.Count;
        return res;
    }
}
'@
}

function Invoke-XllSoak {
    param(
        [string]$Fn,            # registered thread-safe, called as =Fn(ROW(),COLUMN())
        [string]$ArgFormat,     # the args text for one call: {0} = row, {1} = column
        [string]$TypeText,      # the entry row's typetext
        [string]$RetType,       # the exit row's rettype
        [int]$Offset = 0,       # the formula passes ROW()-Offset and COLUMN()-Offset
        [switch]$Natural,       # the function answers arguments below 1 with a message
        [int]$Rows = 1000, [int]$Cols = 20, [int]$Passes = 20, [int]$Threads = 8
    )
    . (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')

    try {
        $sx = Connect-TestExcel
        $app = $sx.App
        Set-XRaySessionDefaults $sx
        $paths = Get-XRayPaths $sx.ProcId

        # Complete-Test exits the process, so automatic calculation is restored before every verdict.
        function Finish([switch]$Pass, [string]$Detail) {
            try { $app.Calculation = -4105 } catch {}
            if ($Pass) { Complete-Test -Pass -Detail $Detail } else { Complete-Test -Fail -Detail $Detail }
        }

        $stem = "Soak$Fn"
        New-XRayMacroBook $sx $stem -Format xlsx -Prepare {
            param($sheet)
            $sheet.Range('A1').Resize($Rows, $Cols).Formula = $(if ($Offset) { "=$Fn(ROW()-$Offset,COLUMN()-$Offset)" } else { "=$Fn(ROW(),COLUMN())" })
        }
        $book = Get-XRayMacroBook $stem
        $ws = $book.Sheet
        $leaf = Split-Path $book.Path -Leaf

        # Manual, so each Calculate does the work; a fixed thread count, so MTC cannot decide one is enough.
        $app.Calculation = -4135
        $app.MultiThreadedCalculation.Enabled = $true
        $app.MultiThreadedCalculation.ThreadMode = 1
        $app.MultiThreadedCalculation.ThreadCount = $Threads

        # PAUSE, not DROP: every row is checked, so none may be lost.
        $echo = Set-XRayTraceParam $sx $null 'BUFFERWHENFULL' 'PAUSE'
        if ($echo -notmatch 'PAUSE') { Finish -Detail "BUFFERWHENFULL PAUSE: $echo" }

        $mark = Get-LogLength $paths.Log
        $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
        if ($pressed -ne 'pressed') { Finish -Detail "arm: $pressed" }
        [void](Wait-LogLine $paths.Log 'VBA tracing: ' $mark)

        $sw = [Diagnostics.Stopwatch]::StartNew()
        for ($p = 1; $p -le $Passes; $p++) {
            $ws.UsedRange.Dirty()
            $app.Calculate()
            if (-not (Wait-XRayCalcDone $app 120)) { Finish -Detail "pass ${p}: calculation did not finish in 120 s" }
        }
        $calcSecs = [math]::Round($sw.Elapsed.TotalSeconds, 1)

        $mark2 = Get-LogLength $paths.Log
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Finish -Detail $lossy }
        [void](Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2 180)

        $problems = @()

        $values = $ws.Range('A1').Resize($Rows, $Cols).Value2
        $wrong = 0; $firstWrong = @(); $messages = 0
        for ($r = 1; $r -le $Rows; $r++) {
            for ($c = 1; $c -le $Cols; $c++) {
                $expect = [XRaySoakCheck]::Expected($r, $c, $Offset, [bool]$Natural)
                if ($expect -notmatch '^-?\d+:-?\d+$') { $messages++ }
                if ([string]$values[$r, $c] -ne $expect) {
                    $wrong++
                    if ($firstWrong.Count -lt 5) { $firstWrong += "R${r}C${c}='$($values[$r, $c])'" }
                }
            }
        }
        if ($wrong) { $problems += "$wrong cell(s) do not hold what their address gives, e.g. $($firstWrong -join ', ')" }

        $sw.Restart()
        $res = [XRaySoakCheck]::Run((Get-XRayTraceCsv $sx.ProcId), $script:TraceHeader, $leaf, $Fn,
                                    $Rows, $Cols, $Passes, $ArgFormat, $TypeText, $RetType, $Offset, [bool]$Natural)
        $checkSecs = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        if ($res.ProblemCount) { $problems += "trace: $($res.ProblemCount) problem(s), first: $($res.Problems -join ' | ')" }
        $want = $Rows * $Cols * $Passes
        if ($res.Entries -ne $want) { $problems += "$($res.Entries) entries, expected $want" }
        if ($res.Threads -lt 2) { $problems += 'every call ran on one thread, so multithreaded calculation never engaged' }

        if ($problems.Count) { Finish -Detail ($problems -join '; ') }
        Finish -Pass -Detail ("{0}: {1} calls on {2} threads, calculated in {3} s; {4} trace rows checked in {5} s; every cell and row matched its address ({6} cells are argument messages)" -f `
                              $Fn, $res.Entries, $res.Threads, $calcSecs, $res.Lines, $checkSecs, $messages)
    }
    catch {
        try { $app.Calculation = -4105 } catch {}
        Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
    }
}
