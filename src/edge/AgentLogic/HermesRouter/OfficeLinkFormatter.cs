namespace HelloWorldA365.AgentLogic.HermesRouter;

using System.Text.RegularExpressions;

/// <summary>
/// Turns Word, Excel and PowerPoint addresses in a Teams reply into Markdown links titled with the
/// document's name. Hermes is told to do this itself and its document tools hand it the exact link;
/// this is the guard for replies that still carry a bare OneDrive or SharePoint address, or a link
/// whose text is the address or a word like "here".
/// </summary>
public static class OfficeLinkFormatter
{
    private static readonly Regex Links = new(
        @"<a\b[^>]*>.*?</a>"
        + @"|\[(?<text>[^\]\n]*)\]\((?<mdurl>https?://[^\s)]+)\)"
        + @"|<(?<angleurl>https?://[^\s>]+)>"
        + @"|(?<bareurl>https?://[^\s<>""'()\[\]]+)",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    private static readonly HashSet<string> OfficeExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".doc", ".docm", ".dotx",
        ".xlsx", ".xls", ".xlsm", ".xlsb",
        ".pptx", ".ppt", ".pptm", ".ppsx",
    };

    private static readonly HashSet<string> GenericLinkText = new(StringComparer.OrdinalIgnoreCase)
    {
        "", "here", "link", "this link", "the link", "click here", "open", "open it", "view",
        "document", "the document", "file", "the file", "workbook", "deck", "presentation",
    };

    public static string Format(string? reply)
    {
        if (string.IsNullOrEmpty(reply) || reply.IndexOf("http", StringComparison.OrdinalIgnoreCase) < 0)
            return reply ?? string.Empty;

        return Links.Replace(reply, match =>
        {
            if (match.Groups["mdurl"].Success)
            {
                var url = match.Groups["mdurl"].Value;
                var text = match.Groups["text"].Value.Trim().TrimEnd('.', ':');
                var title = TitleFor(url);
                if (title == null) return match.Value;
                var generic = GenericLinkText.Contains(text)
                    || text.StartsWith("http", StringComparison.OrdinalIgnoreCase);
                return generic ? $"[{title}]({url})" : match.Value;
            }
            if (match.Groups["angleurl"].Success)
            {
                var url = match.Groups["angleurl"].Value;
                var title = TitleFor(url);
                return title == null ? match.Value : $"[{title}]({url})";
            }
            if (match.Groups["bareurl"].Success)
            {
                var raw = match.Groups["bareurl"].Value;
                var url = raw.TrimEnd('.', ',', ';', ':', '!', '?');
                var title = TitleFor(url);
                return title == null ? match.Value : $"[{title}]({url}){raw[url.Length..]}";
            }
            return match.Value;
        });
    }

    /// <summary>The link text for an Office document address, or null if it isn't one.</summary>
    internal static string? TitleFor(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return null;
        var host = uri.Host.ToLowerInvariant();
        if (!(host.EndsWith(".sharepoint.com") || host == "1drv.ms"
              || host == "onedrive.live.com" || host.EndsWith(".onedrive.live.com")))
            return null;

        // Office for the web: .../_layouts/15/Doc.aspx?sourcedoc={id}&file=Name.docx
        var name = QueryValue(uri.Query, "file");
        if (string.IsNullOrWhiteSpace(name))
        {
            var last = uri.AbsolutePath.TrimEnd('/');
            name = Uri.UnescapeDataString(last[(last.LastIndexOf('/') + 1)..]);
        }
        var extension = Path.GetExtension(name ?? string.Empty);
        if (!string.IsNullOrEmpty(extension) && OfficeExtensions.Contains(extension))
        {
            var stem = Path.GetFileNameWithoutExtension(name)!.Trim();
            if (stem.Length > 0) return Escape(stem);
        }

        // Sharing links carry the app, not the name: /:w:/, /:x:/, /:p:/
        var path = uri.AbsolutePath;
        if (path.StartsWith("/:w:/", StringComparison.OrdinalIgnoreCase)) return "Word document";
        if (path.StartsWith("/:x:/", StringComparison.OrdinalIgnoreCase)) return "Excel workbook";
        if (path.StartsWith("/:p:/", StringComparison.OrdinalIgnoreCase)) return "PowerPoint presentation";
        return null;
    }

    private static string? QueryValue(string query, string key)
    {
        foreach (var pair in query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var at = pair.IndexOf('=');
            if (at <= 0) continue;
            if (!string.Equals(pair[..at], key, StringComparison.OrdinalIgnoreCase)) continue;
            return Uri.UnescapeDataString(pair[(at + 1)..].Replace('+', ' '));
        }
        return null;
    }

    private static string Escape(string title) => title.Replace("[", "\\[").Replace("]", "\\]");
}
