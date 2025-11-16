using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using netDxf;
using netDxf.Collections;
using netDxf.Entities;


internal sealed class MappingColumn
{
    public string? col { get; set; }
    public string? attr { get; set; }
    public string? source { get; set; }
}

internal sealed class Mapping
{
       public List<MappingColumn>? columns { get; set; }
}

internal sealed record TextFragment(double X, double Y, string Value);

internal sealed class Program
{
    private static readonly string[] CutterHeaders =
    {
       
        "رديف",
        "کدسیم",
        "شماره فنی سیم",
        "رنگ سيم",
        "سايزسيم",
        "نوع سيم ",
        "طول برش سيم ",
        "ابتدا",
        "شماره خانه سوکت 1",
        "آماده سازي 1",
        "نام ترمينال 1",
        "طول لختي سيم 1",
        "انتها",
        "شماره خانه سوکت 2",
        "آماده سازي 2",
        "نام ترمينال 2",
        "طول لختي سيم 2",
        "توضيحات"
    };

private const double RowYTolerance = 6.0;


    private static readonly Regex MTextFormatting = new(@"\\[A-Za-z][^;]*;", RegexOptions.Compiled);
private static readonly Regex HasAlphaNumeric = new(@"[\p{L}\p{Nd}]", RegexOptions.Compiled);

public static void Main(string[] args)
{
    Console.OutputEncoding = System.Text.Encoding.UTF8;

    string repoRoot = LocateRepoRoot();
    string dxfPath = args.Length > 0 ? Path.GetFullPath(args[0]) : Path.Combine(repoRoot, "test2-a.dxf");
    string outputPath = args.Length > 1 ? Path.GetFullPath(args[1]) : Path.Combine(repoRoot, "dxf-ext", "output", "test2-a-cutting.xlsx");
    string mappingPath = args.Length > 2 ? Path.GetFullPath(args[2]) : Path.Combine(repoRoot, "dxf-ext", "mapping.json");
    string templatePath = File.Exists(Path.Combine(repoRoot, "cutting2.xlsx")) ? Path.Combine(repoRoot, "cutting2.xlsx") : string.Empty;

    if (!File.Exists(dxfPath))
    {
       
        Console.WriteLine($"❌ DXF file not found: {dxfPath}");
        return;
    }

    Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

    Mapping mapping = LoadMapping(mappingPath);

    DxfDocument document;
    try
    {
        document = DxfDocument.Load(dxfPath);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Unable to load DXF file: {ex.Message}");
        return;
    }

    if (document == null)
    {
        Console.WriteLine("❌ DXF file could not be loaded.");
        return;
    }

    Console.WriteLine($"✅ DXF loaded. Entities: {document.Entities.Count}");

    List<Dictionary<string, string>> rows = ExtractUsingAttributes(document, mapping).ToList();
    if (rows.Count == 0)
    {
        Console.WriteLine("⚠️ Attribute based extraction failed. Falling back to geometric text parsing...");
        rows = ExtractFromText(document).ToList();
    }

    if (rows.Count == 0)
    {
        Console.WriteLine("❌ No usable data found in the DXF file.");
        return;
    }

    WriteWorkbook(rows, outputPath, templatePath);
    Console.WriteLine($"📄 Excel file created: {outputPath}");
}

private static Mapping LoadMapping(string mappingPath)
{
    try
    {
        if (File.Exists(mappingPath))
        {
            string json = File.ReadAllText(mappingPath);
            Mapping? mapping = JsonSerializer.Deserialize<Mapping>(json);
            if (mapping?.columns != null && mapping.columns.Count > 0)
            {
                Console.WriteLine($"ℹ️ Loaded mapping ({mapping.columns.Count} columns).");
                return mapping;
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"⚠️ Unable to load mapping definition: {ex.Message}");
    }

    Console.WriteLine("ℹ️ Mapping file missing or empty. Using default headers only.");
    return new Mapping { columns = new List<MappingColumn>() };
}

private static IEnumerable<Dictionary<string, string>> ExtractUsingAttributes(DxfDocument document, Mapping mapping)
{
    if (mapping.columns == null || mapping.columns.Count == 0)
    {
        return Enumerable.Empty<Dictionary<string, string>>();
    }

    var rows = new List<Dictionary<string, string>>();
    var orderedHeaders = CutterHeaders.ToDictionary(h => h, h => string.Empty, StringComparer.OrdinalIgnoreCase);
    int autoIndex = 1;

    foreach (Insert insert in document.Entities.OfType<Insert>())
    {
        if (insert.Attributes.Count == 0)
        {
            continue;
        }

        var row = new Dictionary<string, string>(orderedHeaders, StringComparer.OrdinalIgnoreCase);
        bool hasValue = false;

        foreach (MappingColumn column in mapping.columns)
        {
            if (string.IsNullOrWhiteSpace(column.col))
            {
                continue;
            }

            if (!row.ContainsKey(column.col))
            {
                row[column.col] = string.Empty;
            }

            if (string.Equals(column.source, "auto_index", StringComparison.OrdinalIgnoreCase))
            {
                row[column.col] = autoIndex.ToString(CultureInfo.InvariantCulture);
                continue;
            }

            if (!string.IsNullOrWhiteSpace(column.attr))
            {
                netDxf.Entities.Attribute? attribute = insert.Attributes
                    .FirstOrDefault(a => string.Equals(a.Tag, column.attr, StringComparison.OrdinalIgnoreCase));

                if (attribute != null)
                {
                    row[column.col] = attribute.Value ?? string.Empty;
                    hasValue = true;
                }
            }
        }

        if (hasValue)
        {
            if (string.IsNullOrWhiteSpace(row[CutterHeaders[0]]))
            {
                row[CutterHeaders[0]] = autoIndex.ToString(CultureInfo.InvariantCulture);
            }

            rows.Add(row);
            autoIndex++;
        }
    }

    if (rows.Count > 0)
    {
        Console.WriteLine($"✅ Extracted {rows.Count} rows from block attributes.");
    }

    return rows;
}

private static IEnumerable<Dictionary<string, string>> ExtractFromText(DxfDocument document)
{
    List<TextFragment> fragments = GetTextFragments(document);
    if (fragments.Count == 0)
    {
        return Enumerable.Empty<Dictionary<string, string>>();
    }

    var orderedFragments = fragments
        .OrderByDescending(f => f.Y)
        .ThenBy(f => f.X)
        .ToList();

    var groupedRows = new List<List<TextFragment>>();
    foreach (TextFragment fragment in orderedFragments)
    {
        if (groupedRows.Count == 0)
        {
            groupedRows.Add(new List<TextFragment> { fragment });
            continue;
        }

        List<TextFragment> currentRow = groupedRows[^1];
        if (Math.Abs(currentRow[0].Y - fragment.Y) <= RowYTolerance)
        {
            currentRow.Add(fragment);
        }
        else
        {
            groupedRows.Add(new List<TextFragment> { fragment });
        }
    }

    int headerRowIndex = FindHeaderRowIndex(groupedRows);
    if (headerRowIndex >= 0)
    {
        groupedRows = groupedRows.Skip(headerRowIndex + 1).ToList();
    }

    var rows = new List<Dictionary<string, string>>();
    int autoIndex = 1;

    foreach (List<TextFragment> group in groupedRows)
    {
        var orderedGroup = group.OrderBy(f => f.X).ToList();
        if (!LooksLikeDataRow(orderedGroup))
        {
            continue;
        }

        var row = CutterHeaders.ToDictionary(h => h, _ => string.Empty);
        for (int col = 0; col < CutterHeaders.Length && col < orderedGroup.Count; col++)
        {
            row[CutterHeaders[col]] = orderedGroup[col].Value;
        }

        if (string.IsNullOrWhiteSpace(row[CutterHeaders[0]]))
        {
            row[CutterHeaders[0]] = autoIndex.ToString(CultureInfo.InvariantCulture);
        }

        rows.Add(row);
        autoIndex++;
    }

    if (rows.Count > 0)
    {
        Console.WriteLine($"✅ Extracted {rows.Count} rows from Text/MText entities.");
    }

    return rows;
}

private static int FindHeaderRowIndex(List<List<TextFragment>> groupedRows)
{
    var normalizedHeaders = CutterHeaders
        .Select(NormalizeHeader)
        .Where(h => !string.IsNullOrWhiteSpace(h))
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    for (int i = 0; i < groupedRows.Count; i++)
    {
        bool containsHeader = groupedRows[i]
            .Select(f => NormalizeHeader(f.Value))
            .Any(normalizedHeaders.Contains);

        if (containsHeader)
        {
            return i;
        }
    }

    return -1;
}

private static bool LooksLikeDataRow(IReadOnlyCollection<TextFragment> fragments)
{
    if (fragments.Count < 2)
    {
        return false;
    }

    int alphanumericFragments = fragments.Count(f => HasAlphaNumeric.IsMatch(f.Value));
    return alphanumericFragments >= 2;
}

private static List<TextFragment> GetTextFragments(DxfDocument document)
{
    var fragments = new List<TextFragment>();

    foreach (Text text in document.Entities.OfType<Text>())
    {
        string value = text.Value?.Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(value))
        {
            fragments.Add(new TextFragment(text.Position.X, text.Position.Y, value));
        }
    }

    foreach (MText mText in document.Entities.OfType<MText>())
    {
        string cleaned = CleanMText(mText.Value);
        if (!string.IsNullOrWhiteSpace(cleaned))
        {
            fragments.Add(new TextFragment(mText.Position.X, mText.Position.Y, cleaned));
        }
    }

    Console.WriteLine($"ℹ️ Collected {fragments.Count} textual fragments.");
    return fragments;
}

private static string CleanMText(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
    {
        return string.Empty;
    }

    string cleaned = value.Trim('{', '}');
    cleaned = MTextFormatting.Replace(cleaned, string.Empty);
    cleaned = cleaned
        .Replace("\\P", " ", StringComparison.Ordinal)
        .Replace("\\p", " ", StringComparison.Ordinal)
        .Replace("\\n", " ", StringComparison.Ordinal)
        .Replace("\\N", " ", StringComparison.Ordinal)
        .Replace("\\~", "~", StringComparison.Ordinal);

    return cleaned.Trim();
}

private static string NormalizeHeader(string? header)
{
    if (string.IsNullOrWhiteSpace(header))
    {
        return string.Empty;
    }

    string normalized = header
        .Replace(" ", string.Empty, StringComparison.Ordinal)
        .Replace("\u200c", string.Empty, StringComparison.Ordinal)
        .Replace("\t", string.Empty, StringComparison.Ordinal)
        .Replace("\r", string.Empty, StringComparison.Ordinal)
        .Replace("\n", string.Empty, StringComparison.Ordinal);

    return normalized.Trim();
}

private static void WriteWorkbook(IEnumerable<Dictionary<string, string>> rows, string outputPath, string templatePath)
{
    XLWorkbook workbook;
    IXLWorksheet worksheet;

    if (!string.IsNullOrWhiteSpace(templatePath) && File.Exists(templatePath))
    {
        workbook = new XLWorkbook(templatePath);
        worksheet = workbook.Worksheet(1);

        var usedRange = worksheet.RangeUsed();
        if (usedRange != null && usedRange.RowCount() > 1)
        {
            int firstDataRow = usedRange.FirstRow().RowNumber() + 1;
            int lastDataRow = usedRange.LastRow().RowNumber();
            if (lastDataRow >= firstDataRow)
            {
                worksheet.Rows(firstDataRow, lastDataRow).Clear(XLClearOptions.Contents);
            }
        }
    }
    else
    {
        workbook = new XLWorkbook();
        worksheet = workbook.AddWorksheet("Sheet1");
        for (int i = 0; i < CutterHeaders.Length; i++)
        {
            worksheet.Cell(1, i + 1).Value = CutterHeaders[i];
        }
    }

    int rowNumber = 2;
    foreach (Dictionary<string, string> row in rows)
    {
        for (int col = 0; col < CutterHeaders.Length; col++)
        {
            string header = CutterHeaders[col];
            string? value = row.TryGetValue(header, out string? val) ? val : string.Empty;
            if (string.IsNullOrWhiteSpace(value) && col == 0)
            {
                value = (rowNumber - 1).ToString(CultureInfo.InvariantCulture);
            }

            worksheet.Cell(rowNumber, col + 1).Value = value;
        }

        rowNumber++;
    }

    workbook.SaveAs(outputPath);
}

private static string LocateRepoRoot()
{
    string current = AppContext.BaseDirectory;
    DirectoryInfo? directory = new DirectoryInfo(current);
    while (directory != null && directory.Exists)
    {
        if (File.Exists(Path.Combine(directory.FullName, "dxf-ext.sln")) ||
            Directory.Exists(Path.Combine(directory.FullName, ".git")))
        {
            return directory.FullName;
        }

        directory = directory.Parent;
    }


    return AppContext.BaseDirectory;
}
}
