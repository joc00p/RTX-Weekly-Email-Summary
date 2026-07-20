using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Reporter;

public record EmailItem(string Subject, string From, string Received, string Body);

public class OutlookService
{
    private const string TargetFolder = "RTX Weekly Team Punch List";

    public Dictionary<string, List<EmailItem>> FetchEmailsByWeek()
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook is not installed or not registered on this machine.");

        dynamic outlook = Activator.CreateInstance(outlookType)!;
        dynamic ns = outlook.GetNamespace("MAPI");

        var diagnostics = new StringBuilder();
        dynamic? folder = null;
        dynamic? items = null;

        try
        {
            foreach (dynamic account in ns.Folders)
            {
                try
                {
                    string accountName = (string)account.Name;
                    diagnostics.AppendLine($"Account: {accountName}");
                    folder = FindFolder(account, TargetFolder, diagnostics, 1);
                    if (folder != null) break;
                }
                catch (System.Exception ex)
                {
                    diagnostics.AppendLine($"  ERROR: {ex.Message}");
                }
            }

            if (folder == null)
                throw new InvalidOperationException(
                    $"Could not find '{TargetFolder}' folder.\n\nFolders scanned:\n{diagnostics}");

            items = folder.Items;
            items.Sort("[ReceivedTime]", true);

            var byWeek = new Dictionary<string, List<EmailItem>>();

            foreach (dynamic msg in items)
            {
                try
                {
                    // 43 = olMailItem
                    if ((int)msg.Class != 43) continue;

                    object rawTime = msg.ReceivedTime;
                    DateTime received = rawTime is double d
                        ? DateTime.FromOADate(d)
                        : Convert.ToDateTime(rawTime);

                    string body = StripQuotedContent((string)msg.Body ?? "");
                    if (string.IsNullOrWhiteSpace(body)) continue;

                    string week = GetWeekLabel(received);
                    if (!byWeek.ContainsKey(week))
                        byWeek[week] = new List<EmailItem>();

                    byWeek[week].Add(new EmailItem(
                        (string)msg.Subject ?? "",
                        (string)msg.SenderName ?? "",
                        received.ToString("ddd MMM dd hh:mm tt"),
                        body.Trim()
                    ));
                }
                catch (System.Runtime.InteropServices.COMException) { }
            }

            return byWeek;
        }
        finally
        {
            if (items != null) try { Marshal.ReleaseComObject(items); } catch { }
            if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            try { Marshal.ReleaseComObject(ns); } catch { }
            try { Marshal.ReleaseComObject(outlook); } catch { }
        }
    }

    public List<string> SearchAddressBook(string query, int maxResults = 100)
    {
        var results = new List<string>();
        if (string.IsNullOrWhiteSpace(query) || query.Length < 2) return results;

        dynamic? outlook = null;
        dynamic? ns = null;
        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Outlook not found.");
            outlook = Activator.CreateInstance(outlookType)!;
            ns = outlook.GetNamespace("MAPI");

            foreach (dynamic addrList in ns.AddressLists)
            {
                try
                {
                    dynamic entries = addrList.AddressEntries;
                    // Try server-side Restrict filter first (fast for GAL)
                    try
                    {
                        string filter = $"[Name] ci_phw '{query}'";
                        dynamic restricted = entries.Restrict(filter);
                        foreach (dynamic entry in restricted)
                        {
                            try
                            {
                                string name = (string)entry.Name;
                                if (!string.IsNullOrWhiteSpace(name))
                                    results.Add(name);
                            }
                            catch { }
                            if (results.Count >= maxResults) return Finish(results);
                        }
                    }
                    catch
                    {
                        // Fallback: iterate manually (safe for small lists like personal Contacts)
                        int count = Math.Min((int)entries.Count, 500);
                        for (int i = 1; i <= count; i++)
                        {
                            try
                            {
                                dynamic entry = entries[i];
                                string name = (string)entry.Name;
                                if (!string.IsNullOrWhiteSpace(name) &&
                                    name.Contains(query, StringComparison.OrdinalIgnoreCase))
                                    results.Add(name);
                            }
                            catch { }
                            if (results.Count >= maxResults) return Finish(results);
                        }
                    }
                }
                catch { }
            }
        }
        catch { }
        finally
        {
            if (ns != null) try { Marshal.ReleaseComObject(ns); } catch { }
            if (outlook != null) try { Marshal.ReleaseComObject(outlook); } catch { }
        }

        return Finish(results);

        static List<string> Finish(List<string> r) =>
            r.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(n => n).ToList();
    }

    private static dynamic? FindFolder(dynamic parent, string name, StringBuilder log, int depth)
    {
        try
        {
            foreach (dynamic sub in parent.Folders)
            {
                try
                {
                    string subName = (string)sub.Name;
                    log.AppendLine($"{new string(' ', depth * 2)}{subName}");
                    if (subName.Contains(name, StringComparison.OrdinalIgnoreCase))
                        return sub;
                    var found = FindFolder(sub, name, log, depth + 1);
                    if (found != null) return found;
                }
                catch { }
            }
        }
        catch { }
        return null;
    }

    private static string StripQuotedContent(string body)
    {
        var lines = body.Split('\n');
        var result = new StringBuilder();

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].TrimEnd();
            var trimmed = line.TrimStart();

            // Stop at common Outlook quote/forward separators only
            if (trimmed.StartsWith("-----Original Message-----", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("________________________________", StringComparison.OrdinalIgnoreCase) ||
                (trimmed.StartsWith("On ") && System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"\bwrote:\s*$")))
                break;

            result.AppendLine(line);
        }

        return result.ToString().Trim();
    }

    private static string GetWeekLabel(DateTime dt)
    {
        int offset = dt.DayOfWeek == DayOfWeek.Sunday ? 6 : (int)dt.DayOfWeek - 1;
        var monday = dt.AddDays(-offset);
        var sunday = monday.AddDays(6);
        return $"Week of {monday:MMM dd} - {sunday:MMM dd, yyyy}";
    }
}
