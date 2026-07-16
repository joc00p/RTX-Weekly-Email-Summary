namespace OutlookSearch.Services;

public record MailFolderNode(string EntryId, string StoreId, string Name, int ItemCount, List<MailFolderNode> Children);

public record SearchOptions(
    string? Keyword,
    string? Subject,
    string? From,
    DateTime? DateFrom,
    DateTime? DateTo);

public class EmailResult
{
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public string Subject { get; init; } = "";
    public string From { get; init; } = "";
    public string To { get; init; } = "";
    public DateTime? ReceivedAt { get; init; }
    public string FolderName { get; init; } = "";
    public string BodyPreview { get; init; } = "";
    public bool HasAttachments { get; init; }
}

/// <summary>
/// Talks to the locally installed Outlook via late-bound COM (the Outlook Object
/// Model). No NuGet packages, no Azure — it attaches to the user's running Outlook
/// profile and can read any mailbox already open in it (own + shared/delegated).
/// </summary>
public class OutlookService : IDisposable
{
    // Outlook enum constants (avoids needing the interop assembly).
    private const int olFolderInbox = 6;
    private const int MaxResultsDefault = 1000;

    private readonly StaTaskScheduler _sta = new();
    private dynamic? _app;
    private dynamic? _ns;

    private void EnsureApp()
    {
        if (_app != null) return;
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Outlook is not installed on this machine (Outlook.Application COM class not found).");
        _app = Activator.CreateInstance(type)!;   // attaches to running Outlook, or starts it
        _ns = _app!.GetNamespace("MAPI");
    }

    /// <summary>Every mailbox / store in the current Outlook profile, as a checkable folder tree.</summary>
    public Task<List<MailFolderNode>> GetMailboxTreesAsync() => _sta.Run(() =>
    {
        EnsureApp();
        var roots = new List<MailFolderNode>();
        dynamic stores = _ns!.Stores;
        int count = stores.Count;
        for (int i = 1; i <= count; i++)
        {
            try
            {
                dynamic store = stores.Item(i);
                string storeId = store.StoreID;
                dynamic root = store.GetRootFolder();
                roots.Add(BuildNode(root, storeId));
            }
            catch { /* skip stores we cannot open */ }
        }
        return roots;
    });

    /// <summary>
    /// Resolve a person by name or email against Outlook's address book (the org GAL / AD)
    /// and return their mailbox folder tree — works when you have delegate/shared access.
    /// </summary>
    public Task<MailFolderNode?> OpenSharedMailboxAsync(string nameOrEmail) => _sta.Run<MailFolderNode?>(() =>
    {
        EnsureApp();
        dynamic recipient = _ns!.CreateRecipient(nameOrEmail);
        recipient.Resolve();
        if (!(bool)recipient.Resolved)
            throw new InvalidOperationException($"Could not resolve '{nameOrEmail}' in the address book.");

        // GetSharedDefaultFolder gives us their Inbox; its parent is the mailbox root,
        // from which we can enumerate the whole (accessible) folder tree.
        dynamic inbox = _ns.GetSharedDefaultFolder(recipient, olFolderInbox);
        dynamic root;
        try { root = inbox.Parent; }
        catch { root = inbox; }

        string storeId = "";
        try { storeId = root.StoreID; } catch { }
        return BuildNode(root, storeId);
    });

    private static MailFolderNode BuildNode(dynamic folder, string storeId)
    {
        var children = new List<MailFolderNode>();
        try
        {
            dynamic subs = folder.Folders;
            int c = subs.Count;
            for (int i = 1; i <= c; i++)
            {
                try { children.Add(BuildNode(subs.Item(i), storeId)); }
                catch { /* skip folders we lack permission to enumerate */ }
            }
        }
        catch { }

        int items = 0;   try { items = folder.Items.Count; } catch { }
        string name = "Folder"; try { name = folder.Name; } catch { }
        string eid = "";  try { eid = folder.EntryID; } catch { }
        string sid = storeId; try { if (string.IsNullOrEmpty(sid)) sid = folder.StoreID; } catch { }

        return new MailFolderNode(eid, sid, name, items, children);
    }

    public Task<List<EmailResult>> SearchAsync(
        IReadOnlyList<(string EntryId, string StoreId, string Name)> folders,
        SearchOptions opts,
        CancellationToken ct) => _sta.Run(() =>
    {
        EnsureApp();
        var dasl = BuildDasl(opts);
        var results = new List<EmailResult>();

        foreach (var (entryId, storeId, folderName) in folders)
        {
            if (ct.IsCancellationRequested) break;
            if (string.IsNullOrEmpty(entryId)) continue;

            dynamic folder;
            try { folder = string.IsNullOrEmpty(storeId) ? _ns!.GetFolderFromID(entryId) : _ns!.GetFolderFromID(entryId, storeId); }
            catch { continue; }

            dynamic items;
            try { items = folder.Items; } catch { continue; }

            // Restrict narrows the set store-side for speed; fall back to the full
            // collection if the DASL query is rejected. Match() below is authoritative.
            dynamic candidates = items;
            if (!string.IsNullOrEmpty(dasl))
            {
                try { candidates = items.Restrict(dasl); } catch { candidates = items; }
            }
            try { candidates.Sort("[ReceivedTime]", true); } catch { }

            int n = 0;
            try { n = candidates.Count; } catch { continue; }

            for (int i = 1; i <= n; i++)
            {
                if (ct.IsCancellationRequested) break;
                if (results.Count >= MaxResultsDefault) break;

                dynamic item;
                try { item = candidates.Item(i); } catch { continue; }

                try
                {
                    string cls = "";
                    try { cls = item.MessageClass; } catch { }
                    if (!cls.StartsWith("IPM.Note")) continue;   // mail items only

                    if (!Match(item, opts)) continue;
                    results.Add(MapItem(item, storeId, folderName));
                }
                catch { /* skip any item that won't map cleanly */ }
            }

            if (results.Count >= MaxResultsDefault) break;
        }

        return results
            .OrderByDescending(r => r.ReceivedAt ?? DateTime.MinValue)
            .ToList();
    });

    public Task<string> GetBodyAsync(string entryId, string storeId) => _sta.Run(() =>
    {
        EnsureApp();
        dynamic item = string.IsNullOrEmpty(storeId)
            ? _ns!.GetItemFromID(entryId)
            : _ns!.GetItemFromID(entryId, storeId);
        try
        {
            string body = item.Body;
            if (!string.IsNullOrWhiteSpace(body)) return body;
        }
        catch { }
        return "";
    });

    // In-memory predicate — the source of truth for whether an item matches.
    private static bool Match(dynamic item, SearchOptions o)
    {
        try
        {
            DateTime? received = TryGetDate(item);
            if (o.DateFrom.HasValue && (received is null || received.Value < o.DateFrom.Value.Date))
                return false;
            if (o.DateTo.HasValue && (received is null || received.Value >= o.DateTo.Value.Date.AddDays(1)))
                return false;

            if (!string.IsNullOrEmpty(o.Subject))
            {
                string subj = SafeStr(() => item.Subject);
                if (subj.IndexOf(o.Subject, StringComparison.OrdinalIgnoreCase) < 0) return false;
            }

            if (!string.IsNullOrEmpty(o.From))
            {
                string sn = SafeStr(() => item.SenderName);
                string se = SafeStr(() => item.SenderEmailAddress);
                if (sn.IndexOf(o.From, StringComparison.OrdinalIgnoreCase) < 0 &&
                    se.IndexOf(o.From, StringComparison.OrdinalIgnoreCase) < 0)
                    return false;
            }

            if (!string.IsNullOrEmpty(o.Keyword))
            {
                string subj = SafeStr(() => item.Subject);
                string body = SafeStr(() => item.Body);
                if (subj.IndexOf(o.Keyword, StringComparison.OrdinalIgnoreCase) < 0 &&
                    body.IndexOf(o.Keyword, StringComparison.OrdinalIgnoreCase) < 0)
                    return false;
            }

            return true;
        }
        catch { return false; }
    }

    private static EmailResult MapItem(dynamic item, string storeId, string folderName)
    {
        string from = SafeStr(() => item.SenderName);
        string fromEmail = SafeStr(() => item.SenderEmailAddress);
        if (!string.IsNullOrEmpty(fromEmail) && !string.Equals(from, fromEmail, StringComparison.OrdinalIgnoreCase))
            from = string.IsNullOrEmpty(from) ? fromEmail : $"{from} <{fromEmail}>";

        string preview = SafeStr(() => item.Body);
        if (preview.Length > 200) preview = preview[..200];
        preview = preview.Replace("\r", " ").Replace("\n", " ").Trim();

        bool hasAtt = false;
        try { hasAtt = item.Attachments.Count > 0; } catch { }

        string sid = storeId;
        if (string.IsNullOrEmpty(sid)) sid = SafeStr(() => item.Parent.StoreID);

        return new EmailResult
        {
            EntryId = SafeStr(() => item.EntryID),
            StoreId = sid,
            Subject = SafeStr(() => item.Subject) is { Length: > 0 } s ? s : "(no subject)",
            From = from,
            To = SafeStr(() => item.To),
            ReceivedAt = TryGetDate(item),
            FolderName = folderName,
            BodyPreview = preview,
            HasAttachments = hasAtt
        };
    }

    private static DateTime? TryGetDate(dynamic item)
    {
        try { return (DateTime)item.ReceivedTime; } catch { }
        try { return (DateTime)item.SentOn; } catch { }
        return null;
    }

    private static string SafeStr(Func<dynamic> get)
    {
        try { return (string)(get() ?? "") ?? ""; }
        catch { return ""; }
    }

    // Builds an Outlook DASL (@SQL) filter to pre-narrow results store-side.
    private static string BuildDasl(SearchOptions o)
    {
        var clauses = new List<string>();

        if (!string.IsNullOrWhiteSpace(o.Subject))
            clauses.Add($"\"urn:schemas:httpmail:subject\" LIKE '%{Esc(o.Subject)}%'");

        if (!string.IsNullOrWhiteSpace(o.From))
            clauses.Add($"(\"urn:schemas:httpmail:fromname\" LIKE '%{Esc(o.From)}%' " +
                        $"OR \"urn:schemas:httpmail:fromemail\" LIKE '%{Esc(o.From)}%')");

        if (o.DateFrom.HasValue)
            clauses.Add($"\"urn:schemas:httpmail:datereceived\" >= '{o.DateFrom.Value:yyyy-MM-dd} 00:00'");

        if (o.DateTo.HasValue)
            clauses.Add($"\"urn:schemas:httpmail:datereceived\" <= '{o.DateTo.Value:yyyy-MM-dd} 23:59'");

        if (!string.IsNullOrWhiteSpace(o.Keyword))
            clauses.Add($"(\"urn:schemas:httpmail:subject\" LIKE '%{Esc(o.Keyword)}%' " +
                        $"OR \"urn:schemas:httpmail:textdescription\" LIKE '%{Esc(o.Keyword)}%')");

        return clauses.Count == 0 ? "" : "@SQL=" + string.Join(" AND ", clauses);
    }

    private static string Esc(string s) => s.Trim().Replace("'", "''");

    public void Dispose() => _sta.Dispose();
}
