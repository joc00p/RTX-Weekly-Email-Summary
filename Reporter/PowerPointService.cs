using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Reporter;

public class PowerPointService
{
    private string _templatePath;
    private readonly TeamConfig _teamConfig;
    private const string NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public PowerPointService(string templatePath, TeamConfig teamConfig)
    {
        _templatePath = templatePath;
        _teamConfig = teamConfig;
    }

    public bool TemplateExists => File.Exists(_templatePath);
    public string TemplatePath { get => _templatePath; set => _templatePath = value; }

    public void Export(string weekLabel, string reportText, TowerMetrics metrics, string outputPath)
    {
        File.Copy(_templatePath, outputPath, overwrite: true);

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var entry = archive.GetEntry("ppt/slides/slide1.xml")
            ?? throw new InvalidOperationException("slide1.xml not found in PPTX.");

        string xmlText;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
            xmlText = reader.ReadToEnd();

        // Update Key Accomplishments via DOM
        var doc = new XmlDocument();
        doc.LoadXml(xmlText);
        var nsm = new XmlNamespaceManager(doc.NameTable);
        nsm.AddNamespace("a", NS);
        nsm.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        UpdateKeyAccomplishments(doc, nsm, reportText);
        UpdateExecutiveSummary(doc, nsm, reportText);
        UpdateManagedServicesMetrics(doc, nsm, metrics);

        // Serialize to MemoryStream as UTF-8 so the XML declaration is correct
        using var ms = new MemoryStream();
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, OmitXmlDeclaration = false };
        using (var xw = XmlWriter.Create(ms, settings))
            doc.Save(xw);

        // Decode to string, replace the date, re-encode as UTF-8
        var serialized = new UTF8Encoding(false).GetString(ms.ToArray());
        var finalXml = UpdateDateInXml(serialized, weekLabel);
        var finalBytes = new UTF8Encoding(false).GetBytes(finalXml);

        entry.Delete();
        var newEntry = archive.CreateEntry("ppt/slides/slide1.xml");
        using var entryStream = newEntry.Open();
        entryStream.Write(finalBytes, 0, finalBytes.Length);
    }

    private static string UpdateDateInXml(string xmlText, string weekLabel)
    {
        // Extract the end date of the LAST selected week (handles multi-week labels)
        // weekLabel: "Week of Jun 23 - Jun 29, 2025" or "Week of ... through Week of Jun 23 - Jun 29, 2025"
        var matches = Regex.Matches(weekLabel, @"-\s+(\w+\.?\s+\d+,?\s*\d{4})");
        if (matches.Count == 0) return xmlText;
        if (!DateTime.TryParse(matches[^1].Groups[1].Value,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AllowWhiteSpaces,
                out var endDate)) return xmlText;
        var dateStr = endDate.ToString("MM/dd/yyyy");

        // Replace "as of XX/XX/XXXX" text run
        xmlText = Regex.Replace(xmlText, @"as of \d{1,2}/\d{1,2}/\d{4}", $"as of {dateStr}",
            RegexOptions.IgnoreCase);
        // Replace standalone date-only text run (the upper-right corner date)
        xmlText = Regex.Replace(xmlText, @"(<a:t>)\d{1,2}/\d{1,2}/\d{4}(</a:t>)", $"${{1}}{dateStr}${{2}}");
        return xmlText;
    }

    private string GetTeam(string personName) => _teamConfig.GetTeam(personName);

    private void UpdateKeyAccomplishments(XmlDocument doc, XmlNamespaceManager nsm, string reportText)
    {
        // Find the table whose first cell header is "Key Accomplishments"
        XmlNode? kaCell = null;
        var tables = doc.SelectNodes("//a:tbl", nsm)!.Cast<XmlNode>();
        foreach (var table in tables)
        {
            var firstCellText = string.Concat(
                table.SelectNodes("a:tr[1]/a:tc[1]//a:t", nsm)!.Cast<XmlNode>().Select(n => n.InnerText));
            if (!firstCellText.Trim().StartsWith("Key Accomplishments", StringComparison.OrdinalIgnoreCase))
                continue;

            var rows = table.SelectNodes("a:tr", nsm)!.Cast<XmlNode>().ToList();
            // Find the content row: first row after the header whose first cell doesn't repeat the header text
            for (int r = 1; r < rows.Count; r++)
            {
                var cellText = string.Concat(
                    rows[r].SelectNodes("a:tc[1]//a:t", nsm)!.Cast<XmlNode>().Select(n => n.InnerText));
                if (!cellText.Trim().StartsWith("Key Accomplishments", StringComparison.OrdinalIgnoreCase))
                {
                    kaCell = rows[r].SelectSingleNode("a:tc[1]", nsm);
                    break;
                }
            }
            break;
        }

        if (kaCell == null) return;

        var txBody = kaCell.SelectSingleNode("a:txBody", nsm)!;
        foreach (var p in txBody.SelectNodes("a:p", nsm)!.Cast<XmlNode>().ToList())
            txBody.RemoveChild(p);

        // Group parsed sections by team. Only the real towers are shown — senders that
        // don't match a tower ("Other") are never rendered under Key Accomplishments.
        var sections = ParseReportSections(reportText);
        var teamOrder = TeamConfig.TowerNames;
        var byTeam = sections.GroupBy(s => GetTeam(s.Name))
                             .ToDictionary(g => g.Key, g => g.ToList());

        foreach (var team in teamOrder)
        {
            if (!byTeam.TryGetValue(team, out var members) || members.Count == 0) continue;

            // Bold team header
            txBody.AppendChild(MakeParagraph(doc, team, bold: true));

            // Round-robin: take 1 bullet from each member in turn, repeat until 5 total
            const int MaxPerTeam = 5;
            var memberBullets = members
                .Select(m => m.Bullets.Where(b => !string.IsNullOrWhiteSpace(b)).ToList())
                .Where(b => b.Count > 0)
                .ToList();
            var allBullets = new List<string>();
            int round = 0;
            while (allBullets.Count < MaxPerTeam)
            {
                bool added = false;
                foreach (var bullets in memberBullets)
                {
                    if (round < bullets.Count)
                    {
                        allBullets.Add(bullets[round]);
                        added = true;
                        if (allBullets.Count >= MaxPerTeam) break;
                    }
                }
                if (!added) break;
                round++;
            }
            foreach (var bullet in allBullets)
                txBody.AppendChild(MakeParagraph(doc, $"• {bullet}", bold: false));

            txBody.AppendChild(MakeEmptyParagraph(doc));
        }
    }

    private static List<(string Name, List<string> Bullets)> ParseReportSections(string text)
    {
        var result = new List<(string, List<string>)>();
        string? currentName = null;
        var currentBullets = new List<string>();

        foreach (var rawLine in text.Split('\n'))
        {
            var line = rawLine.Trim();

            // Stop at the divider or any section heading (Summary, Executive Summary, etc.)
            if (line.StartsWith("---") || line.StartsWith("###") || line.StartsWith("##"))
            {
                if (currentName != null)
                {
                    result.Add((currentName, currentBullets));
                    currentName = null;
                    currentBullets = new List<string>();
                }
                continue;
            }

            if (line.StartsWith("**") && line.EndsWith("**") && line.Length > 4)
            {
                if (currentName != null)
                    result.Add((currentName, currentBullets));
                currentName = line[2..^2];
                currentBullets = new List<string>();
            }
            else if ((line.StartsWith("-") || line.StartsWith("•")) && currentName != null)
            {
                var bullet = line.TrimStart('-', '•').TrimStart();
                if (!string.IsNullOrWhiteSpace(bullet))
                    currentBullets.Add(bullet);
            }
        }
        if (currentName != null)
            result.Add((currentName, currentBullets));

        return result;
    }

    private static XmlElement MakeParagraph(XmlDocument doc, string text, bool bold)
    {
        var p = doc.CreateElement("a", "p", NS);

        var pPr = doc.CreateElement("a", "pPr", NS);
        pPr.SetAttribute("lvl", "0");
        pPr.SetAttribute("algn", "l");
        AppendChild(pPr, "a", "lnSpc", NS, lnSpc =>
            AppendChild(lnSpc, "a", "spcPct", NS, x => x.SetAttribute("val", "100000")));
        AppendChild(pPr, "a", "spcBef", NS, x =>
            AppendChild(x, "a", "spcPts", NS, y => y.SetAttribute("val", "0")));
        AppendChild(pPr, "a", "spcAft", NS, x =>
            AppendChild(x, "a", "spcPts", NS, y => y.SetAttribute("val", "0")));
        pPr.AppendChild(doc.CreateElement("a", "buNone", NS));
        p.AppendChild(pPr);

        var r = doc.CreateElement("a", "r", NS);
        var rPr = doc.CreateElement("a", "rPr", NS);
        rPr.SetAttribute("lang", "en-US");
        rPr.SetAttribute("sz", "800");
        rPr.SetAttribute("b", bold ? "1" : "0");
        rPr.SetAttribute("i", "0");
        rPr.SetAttribute("u", "none");
        rPr.SetAttribute("strike", "noStrike");
        rPr.SetAttribute("kern", "1200");
        rPr.SetAttribute("noProof", "0");
        AppendChild(rPr, "a", "solidFill", NS, fill =>
            AppendChild(fill, "a", "srgbClr", NS, c => c.SetAttribute("val", "000000")));
        rPr.AppendChild(doc.CreateElement("a", "effectLst", NS));
        AppendChild(rPr, "a", "latin", NS, x => x.SetAttribute("typeface", "Aptos"));
        r.AppendChild(rPr);

        var t = doc.CreateElement("a", "t", NS);
        t.InnerText = text;
        r.AppendChild(t);
        p.AppendChild(r);

        return p;
    }

    private static XmlElement MakeEmptyParagraph(XmlDocument doc)
    {
        var p = doc.CreateElement("a", "p", NS);
        var endParaRPr = doc.CreateElement("a", "endParaRPr", NS);
        endParaRPr.SetAttribute("lang", "en-US");
        endParaRPr.SetAttribute("sz", "800");
        endParaRPr.SetAttribute("dirty", "0");
        p.AppendChild(endParaRPr);
        return p;
    }

    private void UpdateExecutiveSummary(XmlDocument doc, XmlNamespaceManager nsm, string reportText)
    {
        var execText = ParseExecutiveSummary(reportText);
        if (string.IsNullOrWhiteSpace(execText)) return;

        XmlNode? contentCell = null;
        foreach (var table in doc.SelectNodes("//a:tbl", nsm)!.Cast<XmlNode>())
        {
            var firstCellText = string.Concat(
                table.SelectNodes("a:tr[1]/a:tc[1]//a:t", nsm)!.Cast<XmlNode>().Select(n => n.InnerText));
            if (!firstCellText.Trim().StartsWith("Executive Summary", StringComparison.OrdinalIgnoreCase))
                continue;
            var rows = table.SelectNodes("a:tr", nsm)!.Cast<XmlNode>().ToList();
            // Find the content row: first row after the header whose first cell doesn't repeat the header text
            for (int r = 1; r < rows.Count; r++)
            {
                var cellText = string.Concat(
                    rows[r].SelectNodes("a:tc[1]//a:t", nsm)!.Cast<XmlNode>().Select(n => n.InnerText));
                if (!cellText.Trim().StartsWith("Executive Summary", StringComparison.OrdinalIgnoreCase))
                {
                    contentCell = rows[r].SelectSingleNode("a:tc[1]", nsm);
                    break;
                }
            }
            break;
        }

        if (contentCell == null) return;

        var txBody = contentCell.SelectSingleNode("a:txBody", nsm)!;
        foreach (var p in txBody.SelectNodes("a:p", nsm)!.Cast<XmlNode>().ToList())
            txBody.RemoveChild(p);

        txBody.AppendChild(MakeParagraph(doc, execText, bold: false));
        txBody.AppendChild(MakeEmptyParagraph(doc));
    }

    private static string ParseExecutiveSummary(string text)
    {
        bool inExec = false;
        var sb = new System.Text.StringBuilder();
        foreach (var rawLine in text.Split('\n'))
        {
            var line = rawLine.Trim();
            if (line.StartsWith("### Executive Summary", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("## Executive Summary", StringComparison.OrdinalIgnoreCase))
            {
                inExec = true;
                continue;
            }
            if (inExec)
            {
                if (line.StartsWith("#")) break;
                if (!string.IsNullOrWhiteSpace(line))
                {
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(line);
                }
            }
        }
        return sb.ToString().Trim();
    }

    // Updates the counts in the lower-right "Managed Services Tasks" table by replacing only
    // the numbers in the existing runs, so all template formatting (bold labels, font, layout)
    // is preserved. A null metric leaves that number as-is.
    private static void UpdateManagedServicesMetrics(XmlDocument doc, XmlNamespaceManager nsm, TowerMetrics m)
    {
        if (m == null || !m.AnyFound) return;

        // Locate the cell by its unique anchor text.
        XmlNode? cell = null;
        foreach (XmlNode t in doc.SelectNodes("//a:tc//a:t", nsm)!)
        {
            if (t.InnerText.Contains("SQL Databases") || t.InnerText.Contains("Servers in Sandbox"))
            {
                cell = t.SelectSingleNode("ancestor::a:tc", nsm);
                break;
            }
        }
        if (cell == null) return;

        foreach (XmlNode t in cell.SelectNodes(".//a:t", nsm)!)
        {
            var s = t.InnerText;
            // "20 instances with 104 RISE servers, 4 live apps"
            s = ReplaceMetric(s, @"(\d+)(\s+instances with\s+)(\d+)(\s+RISE servers,\s+)(\d+)(\s+live apps)",
                m.SapInstances, m.SapRiseServers, m.SapRiseLiveApps);
            // "19 XETA servers, 4 live apps"
            s = ReplaceMetric(s, @"(\d+)(\s+XETA servers,\s+)(\d+)(\s+live apps)",
                m.SapXetaServers, m.SapXetaLiveApps);
            // "197 SQL Databases"
            s = ReplaceMetric(s, @"(\d+)(\s+SQL Databases)", m.SqlDatabases);
            // "113 Servers in Sandbox, DEV, PROD"
            s = ReplaceMetric(s, @"(\d+)(\s+Servers in Sandbox)", m.CloudServers);
            if (s != t.InnerText) t.InnerText = s;
        }
    }

    // Replaces the odd-numbered capture groups (the numbers) with the supplied values,
    // keeping the even groups (the literal separators). A null value keeps the original number.
    private static string ReplaceMetric(string input, string pattern, params int?[] values)
    {
        return Regex.Replace(input, pattern, match =>
        {
            var sb = new StringBuilder();
            int valIdx = 0;
            for (int g = 1; g < match.Groups.Count; g++)
            {
                if (g % 2 == 1) // number slot
                {
                    var v = valIdx < values.Length ? values[valIdx] : null;
                    sb.Append(v?.ToString() ?? match.Groups[g].Value);
                    valIdx++;
                }
                else // literal separator
                {
                    sb.Append(match.Groups[g].Value);
                }
            }
            return sb.ToString();
        });
    }

    private static XmlElement AppendChild(XmlNode parent, string prefix, string localName, string ns,
        Action<XmlElement>? configure = null)
    {
        var el = ((XmlDocument)(parent.OwnerDocument ?? (XmlDocument)parent))
            .CreateElement(prefix, localName, ns);
        configure?.Invoke(el);
        parent.AppendChild(el);
        return el;
    }
}
