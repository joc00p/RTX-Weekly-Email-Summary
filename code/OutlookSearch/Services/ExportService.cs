using System.Runtime.InteropServices;
using System.Text;

namespace OutlookSearch.Services;

/// <summary>One email's fields as they will be written to the export file.</summary>
public record ExportRow(string Received, string From, string To, string Subject, string Folder, string Body);

/// <summary>
/// Exports a set of emails to a single file. Plain text is written directly;
/// .xls/.xlsx, .doc/.docx, .ppt/.pptx and .pdf are produced via late-bound Office
/// COM automation (Excel / Word / PowerPoint), so no NuGet packages are needed.
/// All Office work runs on a dedicated STA thread.
/// </summary>
public class ExportService : IDisposable
{
    private readonly StaTaskScheduler _sta = new();

    // Office SaveAs format codes.
    private const int WdFormatDocument97 = 0;    // .doc
    private const int WdFormatDocumentDefault = 16; // .docx
    private const int WdFormatPdf = 17;          // .pdf (via Word)
    private const int XlExcel8 = 56;             // .xls
    private const int XlOpenXmlWorkbook = 51;    // .xlsx
    private const int PpSaveAsPresentation = 1;  // .ppt
    private const int PpSaveAsOpenXml = 24;      // .pptx
    private const int PpLayoutText = 2;
    private const int MsoFalse = 0;

    private const int ExcelCellMax = 32000;      // Excel hard limit is 32767 chars
    private const int SlideBodyMax = 4000;       // keep slides legible

    /// <summary>Human-friendly SaveFileDialog filter covering every supported format.</summary>
    public const string FileFilter =
        "Text file (*.txt)|*.txt|" +
        "PDF (*.pdf)|*.pdf|" +
        "Excel workbook (*.xlsx)|*.xlsx|" +
        "Excel 97-2003 (*.xls)|*.xls|" +
        "Word document (*.docx)|*.docx|" +
        "Word 97-2003 (*.doc)|*.doc|" +
        "PowerPoint (*.pptx)|*.pptx|" +
        "PowerPoint 97-2003 (*.ppt)|*.ppt";

    public Task ExportAsync(string path, IReadOnlyList<ExportRow> rows, CancellationToken ct) => _sta.Run(() =>
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".txt":                 ExportTxt(path, rows); break;
            case ".xls": case ".xlsx":   ExportExcel(path, rows, ext); break;
            case ".doc": case ".docx":
            case ".pdf":                 ExportWord(path, rows, ext); break;  // PDF is produced through Word
            case ".ppt": case ".pptx":   ExportPowerPoint(path, rows, ext); break;
            default:                     ExportTxt(path, rows); break;
        }
    });

    // ── Plain text ────────────────────────────────────────────────
    private static void ExportTxt(string path, IReadOnlyList<ExportRow> rows)
    {
        var sb = new StringBuilder();
        foreach (var r in rows)
        {
            sb.AppendLine($"Subject:  {r.Subject}");
            sb.AppendLine($"From:     {r.From}");
            sb.AppendLine($"To:       {r.To}");
            sb.AppendLine($"Received: {r.Received}");
            sb.AppendLine($"Folder:   {r.Folder}");
            sb.AppendLine(new string('-', 90));
            sb.AppendLine(r.Body);
            sb.AppendLine();
            sb.AppendLine(new string('=', 90));
            sb.AppendLine();
        }
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
    }

    // ── Excel (.xlsx / .xls) ──────────────────────────────────────
    private static void ExportExcel(string path, IReadOnlyList<ExportRow> rows, string ext)
    {
        dynamic? app = null, wb = null, ws = null;
        try
        {
            app = CreateApp("Excel.Application");
            app.Visible = false;
            app.DisplayAlerts = false;
            wb = app.Workbooks.Add();
            ws = wb.Worksheets[1];

            int n = rows.Count;
            var data = new object[n + 1, 6];
            data[0, 0] = "Received"; data[0, 1] = "From"; data[0, 2] = "To";
            data[0, 3] = "Subject";  data[0, 4] = "Folder"; data[0, 5] = "Body";
            for (int i = 0; i < n; i++)
            {
                var r = rows[i];
                data[i + 1, 0] = r.Received;
                data[i + 1, 1] = r.From;
                data[i + 1, 2] = r.To;
                data[i + 1, 3] = r.Subject;
                data[i + 1, 4] = r.Folder;
                data[i + 1, 5] = Trunc(r.Body, ExcelCellMax);
            }

            ws.Range[$"A1:F{n + 1}"].Value2 = data;
            ws.Range["A1:F1"].Font.Bold = true;
            ws.Columns.AutoFit();
            ws.Range["F:F"].ColumnWidth = 90;

            DeleteIfExists(path);
            wb.SaveAs(path, ext == ".xls" ? XlExcel8 : XlOpenXmlWorkbook);
            wb.Close(false);
            app.Quit();
        }
        finally { Rel(ws); Rel(wb); Rel(app); }
    }

    // ── Word (.docx / .doc) and PDF ───────────────────────────────
    private static void ExportWord(string path, IReadOnlyList<ExportRow> rows, string ext)
    {
        dynamic? app = null, doc = null;
        try
        {
            app = CreateApp("Word.Application");
            app.Visible = false;
            try { app.DisplayAlerts = 0; } catch { }
            doc = app.Documents.Add();
            dynamic sel = app.Selection;

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sel.Font.Bold = true; sel.Font.Size = 12;
                sel.TypeText(r.Subject); sel.TypeParagraph();
                sel.Font.Bold = false; sel.Font.Size = 10;
                sel.TypeText($"From: {r.From}"); sel.TypeParagraph();
                sel.TypeText($"To: {r.To}"); sel.TypeParagraph();
                sel.TypeText($"Received: {r.Received}    Folder: {r.Folder}"); sel.TypeParagraph();
                sel.TypeParagraph();
                sel.TypeText(r.Body ?? ""); sel.TypeParagraph();
                if (i < rows.Count - 1) sel.InsertNewPage();
            }

            int fmt = ext switch { ".pdf" => WdFormatPdf, ".doc" => WdFormatDocument97, _ => WdFormatDocumentDefault };
            DeleteIfExists(path);
            try { doc.SaveAs2(path, fmt); } catch { doc.SaveAs(path, fmt); }
            doc.Close(false);
            app.Quit();
        }
        finally { Rel(doc); Rel(app); }
    }

    // ── PowerPoint (.pptx / .ppt) ─────────────────────────────────
    private static void ExportPowerPoint(string path, IReadOnlyList<ExportRow> rows, string ext)
    {
        dynamic? app = null, pres = null;
        try
        {
            app = CreateApp("PowerPoint.Application");
            // Prefer a windowless presentation; fall back to a normal one if unsupported.
            try { pres = app.Presentations.Add(MsoFalse); } catch { pres = app.Presentations.Add(); }

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                dynamic slide = pres.Slides.Add(i + 1, PpLayoutText);
                try { slide.Shapes.Title.TextFrame.TextRange.Text = r.Subject; } catch { }
                var content = $"From: {r.From}\rTo: {r.To}\rReceived: {r.Received}\rFolder: {r.Folder}\r\r{Trunc(r.Body, SlideBodyMax)}";
                try { slide.Shapes[2].TextFrame.TextRange.Text = content; } catch { }
                Rel(slide);
            }

            DeleteIfExists(path);
            pres.SaveAs(path, ext == ".ppt" ? PpSaveAsPresentation : PpSaveAsOpenXml);
            pres.Close();
            app.Quit();
        }
        finally { Rel(pres); Rel(app); }
    }

    // ── Helpers ───────────────────────────────────────────────────
    private static dynamic CreateApp(string progId)
    {
        var t = Type.GetTypeFromProgID(progId)
            ?? throw new InvalidOperationException(
                $"{progId.Split('.')[0]} is not installed, so this format can't be exported. Try .txt instead.");
        return Activator.CreateInstance(t)!;
    }

    private static string Trunc(string? s, int max)
    {
        s ??= "";
        return s.Length <= max ? s : s[..max] + "…";
    }

    private static void DeleteIfExists(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

    private static void Rel(object? o)
    {
        try { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }
        catch { }
    }

    public void Dispose() => _sta.Dispose();
}
