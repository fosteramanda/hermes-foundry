namespace HelloWorldA365.AgentLogic.HermesRouter;

using System.Globalization;

/// <summary>
/// Names a Teams Plan step from the tool Hermes is running and its arguments: "-ing" while the
/// step runs ("Searching the web for this week's biggest news"), past tense once it is done
/// ("Searched the web for this week's biggest news"). Names are full and specific, fit the width
/// Teams gives a row without "...", and never repeat within one Plan.
/// </summary>
public sealed class StepNamer
{
    /// <summary>
    /// Longest name Teams shows without cutting it: the collapsed work line is about 210 pixels
    /// wide in Teams web and cut "Searching the web for "What ar..." at 30 characters (21 Nov).
    /// </summary>
    public static int MaxLength { get; set; } = 26;

    private readonly HashSet<string> _used = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, int> _counts = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Tools that are bookkeeping, not work the person would recognise as a step.</summary>
    public static bool IsHidden(string tool)
    {
        var name = Suffix(tool);
        return name is "start_task" or "skill_view" or "skills_list" or "tool_search" or "tool_describe";
    }

    public static string Suffix(string tool)
    {
        if (string.IsNullOrWhiteSpace(tool)) return "working";
        var marker = tool.LastIndexOf("__", StringComparison.Ordinal);
        return (marker >= 0 ? tool[(marker + 2)..] : tool).ToLowerInvariant();
    }

    /// <summary>The running and done names for a new step, unique within this Plan.</summary>
    public (string Running, string Done) Name(string tool, IReadOnlyDictionary<string, string> args)
    {
        args ??= new Dictionary<string, string>();
        var name = Suffix(tool);
        var (running, done) = Base(name, args);
        var key = name + "|" + running;
        var n = _counts.TryGetValue(key, out var c) ? c + 1 : 1;
        string suffix = "", r, d;
        while (true)
        {
            suffix = n > 1 ? " " + Ordinal(n) : "";
            r = Fit(running, MaxLength - suffix.Length) + suffix;
            d = Fit(done, MaxLength - suffix.Length) + suffix;
            if (!_used.Contains(r) || n > 6) break;
            n++;
        }
        _counts[key] = n;
        _used.Add(r);
        return (r, d);
    }

    private static (string, string) Base(string tool, IReadOnlyDictionary<string, string> a)
    {
        string Arg(params string[] keys)
        {
            foreach (var k in keys)
                if (a.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
                    return Clean(v);
            return "";
        }
        (string, string) With(string ing, string ed, string detail) =>
            string.IsNullOrWhiteSpace(detail) ? (ing, ed) : ($"{ing} {detail}", $"{ed} {detail}");

        var query = Arg("query", "q", "search", "keywords");
        var url = Arg("url", "link");
        var file = FileName(Arg("path", "name", "file_name", "filename", "source_path"));
        switch (tool)
        {
            case "web_search": return With("Searching for", "Searched for", Or(Keywords(query), "web results"));
            case "fetch_url":
            case "browse_url": return With("Reading", "Read", Site(url));
            case "screenshot_url": return With("Capturing", "Captured", Site(url));
            case "search_files": return With("Searching OneDrive for", "Searched OneDrive for", Keywords(query));
            case "list_files": return ("Listing OneDrive files", "Listed OneDrive files");
            case "read_file": return With("Reading", "Read", Or(file, "a OneDrive file"));
            case "list_shared_with_me": return ("Checking files shared with me", "Checked files shared with me");
            case "search_emails": return With("Searching email for", "Searched email for", Keywords(query));
            case "list_inbox": return ("Checking the inbox", "Checked the inbox");
            case "read_email": return ("Reading an email", "Read an email");
            case "send_email": return With("Emailing", "Emailed", Person(Arg("to", "recipients")));
            case "reply_email": return ("Replying to the email", "Replied to the email");
            case "send_teams_message": return With("Messaging", "Messaged", Person(Arg("to")));
            case "list_events": return ("Checking the calendar", "Checked the calendar");
            case "create_event": return With("Booking", "Booked", Or(Arg("subject", "title"), "the meeting"));
            case "update_event": return ("Updating the meeting", "Updated the meeting");
            case "invite_to_event": return ("Inviting people to the meeting", "Invited people to the meeting");
            case "attach_to_event": return With("Attaching", "Attached", Or(file, "a file to the meeting"));
            case "create_docx":
            case "create_xlsx":
            case "create_pptx":
            case "create_pdf":
            case "create_file": return With("Creating", "Created", Or(file, Kind(tool)));
            case "edit_docx": return With("Editing", "Edited", Or(file, "the Word document"));
            case "create_folder": return With("Creating folder", "Created folder", Arg("name"));
            case "share_file": return With("Sharing the file with", "Shared the file with", Or(Person(Arg("email", "to")), "a colleague"));
            case "resolve_person": return With("Looking up", "Looked up", Arg("query"));
            case "run_code":
            case "execute_code": return ("Running Python", "Ran Python");
            case "remember": return ("Saving to memory", "Saved to memory");
            case "recall": return ("Checking memory", "Checked memory");
            case "create_routine": return With("Setting up routine", "Set up routine", Arg("name"));
            case "list_routines": return ("Checking routines", "Checked routines");
            case "remove_routine": return With("Removing routine", "Removed routine", Arg("name"));
            case "propose_access_change": return ("Preparing an access request", "Prepared an access request");
            case "list_access_permissions": return ("Checking approved access", "Checked approved access");
            case "state_sync_config": return ("Checking saved-state settings", "Checked saved-state settings");
            case "whoami": return ("Checking my account", "Checked my account");
            case "graph_query": return ("Reading Microsoft 365 data", "Read Microsoft 365 data");
            case "graph_write": return ("Updating Microsoft 365 data", "Updated Microsoft 365 data");
            default:
                var words = tool.Replace('_', ' ').Trim();
                return ($"Running {words}", $"Ran {words}");
        }
    }

    private static string Kind(string tool) => tool switch
    {
        "create_docx" => "a Word document",
        "create_xlsx" => "a spreadsheet",
        "create_pptx" => "a presentation",
        "create_pdf" => "a PDF",
        _ => "a file",
    };

    private static string Ordinal(int n) => n switch
    {
        2 => "again",
        3 => "a third time",
        4 => "a fourth time",
        5 => "a fifth time",
        _ => "once more",
    };

    /// <summary>Shorten the variable part at a word boundary until the whole name fits.</summary>
    private static string Fit(string name, int max)
    {
        name = Clean(name);
        if (name.Length <= max) return name;
        var words = name.Split(' ').ToList();
        while (words.Count > 2 && string.Join(' ', words).Length > max)
            words.RemoveAt(words.Count - 1);
        var fitted = string.Join(' ', words).TrimEnd(',', ';', ':', '-', '"', '\'');
        if (fitted.Count(ch => ch == '"') % 2 == 1)
            fitted += "\"";
        return fitted.Length <= max ? fitted : fitted[..max].TrimEnd();
    }

    private static string Quote(string v) => string.IsNullOrWhiteSpace(v) ? "" : $"\"{v.Trim('"', '\'', ' ')}\"";

    private static string Or(string v, string fallback) => string.IsNullOrWhiteSpace(v) ? fallback : v;

    private static string Person(string v)
    {
        v = Clean(v);
        if (string.IsNullOrWhiteSpace(v)) return "";
        var first = v.Split(',', ';')[0].Trim();
        if (first.Equals("owner", StringComparison.OrdinalIgnoreCase)) return "my manager";
        var at = first.IndexOf('@');
        if (at <= 0) return first;
        var local = first[..at].Replace('.', ' ').Replace('_', ' ');
        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(local);
    }

    private static string FileName(string path)
    {
        path = Clean(path);
        if (string.IsNullOrWhiteSpace(path)) return "";
        var slash = Math.Max(path.LastIndexOf('/'), path.LastIndexOf('\\'));
        return slash >= 0 ? path[(slash + 1)..] : path;
    }

    private static string Site(string url)
    {
        if (!Uri.TryCreate(Clean(url), UriKind.Absolute, out var uri)) return "a web page";
        return uri.Host.StartsWith("www.", StringComparison.OrdinalIgnoreCase) ? uri.Host[4..] : uri.Host;
    }

    private static readonly HashSet<string> Stop = new(StringComparer.OrdinalIgnoreCase)
    {
        "a", "an", "and", "are", "about", "as", "at", "be", "by", "can", "could", "do", "does", "find",
        "for", "from", "get", "give", "how", "i", "in", "is", "it", "me", "my", "of", "on", "or",
        "please", "show", "tell", "that", "the", "this", "to", "was", "were", "what", "whats", "when",
        "where", "which", "who", "why", "will", "with", "you", "your",
    };

    /// <summary>The words of a search that say what it is about, in order, without filler.</summary>
    private static string Keywords(string query)
    {
        var words = Clean(query).Trim('"', '\'', '?', '.', '!').Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(w => w.Trim('"', '\'', '?', '.', ',', '!', ':', ';', '(', ')'))
            .Where(w => w.Length > 0 && !Stop.Contains(w.Replace("'", "")))
            .ToList();
        return string.Join(' ', words);
    }

    private static string Clean(string v) =>
        string.Join(' ', (v ?? "").Replace('\r', ' ').Replace('\n', ' ').Split(' ', StringSplitOptions.RemoveEmptyEntries));
}
