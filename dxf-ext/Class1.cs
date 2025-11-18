using netDxf;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace dxf_ext
{
    public class Class1
    {
        private static readonly Regex NumericCodePattern = new("^[0-9]{3,4}$", RegexOptions.Compiled);
        private static readonly Regex AlphanumericCodePattern = new("^[A-Z][0-9]{2,3}[A-Z]?$", RegexOptions.Compiled);
        public static void Extract()
        { 
        try
        {
            var inputPath = @"c:\tmp\test2-a.dxf";
                var outputPath = @"c:\tmp\wire_codes.xlsx";

                Console.WriteLine($"Loading DXF file: {inputPath}");
            var doc = DxfDocument.Load(inputPath);

        var candidates = CollectTextLikeValues(doc);
        var codes = ExtractWireCodes(candidates)
            .Distinct()
            .OrderBy(code => code, StringComparer.Ordinal)
            .ToList();

            if (codes.Count == 0)
            {
                Console.WriteLine("No wire codes matching the expected pattern were found.");
                
            }

    WriteOutput(outputPath, codes);

    Console.WriteLine("Wire codes detected (sorted):");
            foreach (var code in codes)
            {
                Console.WriteLine(code);
            }

Console.WriteLine($"\nSaved {codes.Count} unique codes to {outputPath}");

        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);

        }
    }

    private static string ResolveInputPath(IReadOnlyList<string> args)
{
    if (args.Count > 0)
    {
        return Path.GetFullPath(args[0]);
    }

    var defaultPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "test2-a.dxf");
    if (!File.Exists(defaultPath))
    {
        throw new FileNotFoundException("DXF file not found. Provide the path as the first argument.");
    }

    return defaultPath;
}

private static string ResolveOutputPath(IReadOnlyList<string> args, string inputPath)
{
    if (args.Count > 1)
    {
        return Path.GetFullPath(args[1]);
    }

    var directory = Path.GetDirectoryName(inputPath)!;
    var fileName = Path.GetFileNameWithoutExtension(inputPath) + "-wire-codes.txt";
    return Path.Combine(directory, fileName);
}

private static IReadOnlyCollection<string> CollectTextLikeValues(DxfDocument doc)
{
    var buffer = new List<string>();

    void Add(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            buffer.Add(value);
        }
    }

    foreach (var text in doc.Entities.Texts)
    {
        Add(text.Value);
    }

    foreach (var mText in doc.Entities.MTexts)
    {
        Add(mText.Value);
    }

    foreach (var insert in doc.Entities.Inserts)
    {
        foreach (var attribute in insert.Attributes)
        {
            Add(attribute.Value);
        }
    }

    return buffer;
}

private static IEnumerable<string> ExtractWireCodes(IEnumerable<string> candidates)
{
    foreach (var candidate in candidates)
    {
        foreach (var token in Tokenize(candidate))
        {
            var normalized = token.ToUpperInvariant();

            if (NumericCodePattern.IsMatch(normalized) || AlphanumericCodePattern.IsMatch(normalized))
            {
                yield return normalized;
            }
        }
    }
}

private static IEnumerable<string> Tokenize(string candidate)
{
    var builder = new StringBuilder();

    foreach (var ch in candidate)
    {
        if (char.IsLetterOrDigit(ch))
        {
            builder.Append(ch);
        }
        else if (builder.Length > 0)
        {
            yield return builder.ToString();
            builder.Clear();
        }
    }

    if (builder.Length > 0)
    {
        yield return builder.ToString();
    }
}

private static void WriteOutput(string outputPath, IReadOnlyCollection<string> codes)
{
    Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
    File.WriteAllLines(outputPath, codes);
}
    }
}
