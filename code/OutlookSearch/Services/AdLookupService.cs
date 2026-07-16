using System.DirectoryServices;
using System.Text;

namespace OutlookSearch.Services;

public record AdPerson(string DisplayName, string Email)
{
    public override string ToString() =>
        string.IsNullOrEmpty(Email) ? DisplayName : $"{DisplayName}  <{Email}>";
}

/// <summary>
/// Looks up people in on-premises Active Directory over LDAP (no cloud, no keys).
/// Uses AD's ANR (Ambiguous Name Resolution) so a single term matches display
/// name, first/last name, alias, and email — the same behavior as Outlook's
/// "check names". All queries run on a background thread.
/// </summary>
public class AdLookupService
{
    private string? _rootPath;
    private bool _resolved;
    private readonly object _lock = new();

    // Resolved lazily on a background thread so a slow/unavailable DC never blocks the UI.
    private string? RootPath()
    {
        lock (_lock)
        {
            if (_resolved) return _rootPath;
            _resolved = true;
            try
            {
                using var rootDse = new DirectoryEntry("LDAP://RootDSE");
                if (rootDse.Properties["defaultNamingContext"].Value is string dnc && dnc.Length > 0)
                    _rootPath = "LDAP://" + dnc;
            }
            catch { _rootPath = null; }
            return _rootPath;
        }
    }

    public Task<List<AdPerson>> SearchAsync(string term, int max, CancellationToken ct) =>
        Task.Run(() => Search(term, max, ct), ct);

    private List<AdPerson> Search(string term, int max, CancellationToken ct)
    {
        var people = new List<AdPerson>();
        var root = RootPath();
        if (root is null || string.IsNullOrWhiteSpace(term)) return people;

        try
        {
            using var entry = new DirectoryEntry(root);
            using var searcher = new DirectorySearcher(entry)
            {
                Filter = $"(&(objectCategory=person)(objectClass=user)(mail=*)(anr={EscapeAnr(term)}))",
                SizeLimit = max,
                ClientTimeout = TimeSpan.FromSeconds(5),
                ServerPageTimeLimit = TimeSpan.FromSeconds(5)
            };
            searcher.PropertiesToLoad.Add("displayName");
            searcher.PropertiesToLoad.Add("mail");

            using var results = searcher.FindAll();
            foreach (SearchResult r in results)
            {
                ct.ThrowIfCancellationRequested();
                var mail = Prop(r, "mail");
                if (mail.Length == 0) continue;
                var name = Prop(r, "displayName");
                people.Add(new AdPerson(name.Length > 0 ? name : mail, mail));
            }
        }
        catch (OperationCanceledException) { throw; }
        catch { /* directory unreachable or transient — return whatever we gathered */ }

        return people
            .OrderBy(p => p.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string Prop(SearchResult r, string name) =>
        r.Properties.Contains(name) && r.Properties[name].Count > 0
            ? r.Properties[name][0]?.ToString() ?? ""
            : "";

    // Escape LDAP filter special characters (RFC 4515) to keep the query safe.
    private static string EscapeAnr(string s)
    {
        var sb = new StringBuilder();
        foreach (char c in s.Trim())
        {
            switch (c)
            {
                case '\\': sb.Append("\\5c"); break;
                case '*':  sb.Append("\\2a"); break;
                case '(':  sb.Append("\\28"); break;
                case ')':  sb.Append("\\29"); break;
                case '\0': sb.Append("\\00"); break;
                default:   sb.Append(c); break;
            }
        }
        return sb.ToString();
    }
}
