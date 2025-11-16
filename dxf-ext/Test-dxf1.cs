using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Text.Json;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using netDxf;
using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions; // Added for pattern matching

class MappingColumn { public string col { get; set; } public string attr { get; set; } public string source { get; set; } }
class Mapping { public List<MappingColumn> columns { get; set; } }

class Program
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string dxfPath = "C:\\tmp\\sample-dxf.dxf";
        string mappingPath = "C:\\Users\\Digibod.ir\\source\\repos\\dxf-ext\\dxf-ext\\mapping.json";

        if (!File.Exists(dxfPath)) { Console.WriteLine($"ERROR: DXF file not found: {dxfPath}"); return; }
        if (!File.Exists(mappingPath)) { Console.WriteLine($"ERROR: mapping.json not found: {mappingPath}"); return; }

        Mapping mapping = JsonSerializer.Deserialize<Mapping>(File.ReadAllText(mappingPath));

        // Load DXF
        DxfDocument dxf = null;
        try
        {
            dxf = DxfDocument.Load(dxfPath);
            Console.WriteLine($"✅ DXF loaded. Entities count: {dxf.Entities.All.Count()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR loading DXF: {ex.Message}");
            return;
        }

        // Utility: get enumerable from a possible property or method using reflection
        IEnumerable SafeGetEnumerable(object owner, string memberName)
        {
            // ... (Your existing SafeGetEnumerable method remains the same) ...
            if (owner == null) return Enumerable.Empty<object>();

            var t = owner.GetType();

            // Try property
            var prop = t.GetProperty(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (prop != null)
            {
                var val = prop.GetValue(owner);
                if (val is IEnumerable e) return e.Cast<object>();
            }

            // Try method (no parameters)
            var m = t.GetMethod(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase, null, Type.EmptyTypes, null);
            if (m != null)
            {
                var val = m.Invoke(owner, null);
                if (val is IEnumerable e2) return e2.Cast<object>();
            }

            // Try field (rare)
            var field = t.GetField(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (field != null)
            {
                var val = field.GetValue(owner);
                if (val is IEnumerable e3) return e3.Cast<object>();
            }

            return Enumerable.Empty<object>();
        }

        var entitiesEnum = SafeGetEnumerable(dxf, "Entities") ?? Enumerable.Empty<object>();
        int entitiesCount = entitiesEnum.Cast<object>().Count();
        Console.WriteLine($"DXF loaded. Entities count: {entitiesCount}");

        var table = new DataTable("Wires");
        foreach (var mc in mapping.columns) table.Columns.Add(mc.col);

        int idx = 1;
        bool dataExtractedFromInserts = false;

        // Try extracting from Inserts/Attributes first (best method)
        var inserts = dxf.Entities.All.OfType<netDxf.Entities.Insert>().ToList();

        if (inserts.Count > 0)
        {
            // The existing Insert/Attribute extraction logic (commented out in your original code)
            // should be uncommented and validated here. For now, we assume it's still bypassed
            // or the Attributes are empty, leading to the fallback.

            // Re-enabling the core logic for Insert extraction:
            foreach (var insObj in inserts)
            {
                var attributes = insObj.Attributes.ToList(); // Insert.Attributes has netDxf.Entities.Attribute instances (values)

                // You were printing the attribute tags/values for debugging, which is good.

                if (attributes.Count > 0)
                {
                    dataExtractedFromInserts = true;
                    var row = table.NewRow();

                    // Initialize row with empty strings
                    foreach (DataColumn column in table.Columns) row[column.ColumnName] = "";

                    foreach (var mc in mapping.columns)
                    {
                        if (mc.source == "auto_index")
                        {
                            row[mc.col] = idx;
                            continue;
                        }
                        if (!string.IsNullOrEmpty(mc.attr))
                        {
                            // Find the attribute by Tag (case-insensitive)
                            var attribute = attributes.FirstOrDefault(a =>
                                string.Equals(a.Tag, mc.attr, StringComparison.OrdinalIgnoreCase)
                            );

                            if (attribute != null)
                            {
                                row[mc.col] = attribute.Value;
                            }
                        }
                    }
                    table.Rows.Add(row);
                    idx++;
                }
            }
            if (dataExtractedFromInserts) Console.WriteLine($"✅ Successfully extracted data from {idx - 1} Insert entities.");
        }


        // Fallback Logic: Parsing Text/MText entities
        if (!dataExtractedFromInserts)
        {
            Console.WriteLine("⚠️ No structured attribute data found. Falling back to parsing Text/MText entities.");

            // --- Revised Parsing Definitions ---
            var singleCharColors = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "R", "B", "Y", "L", "W", "G", "V", "O", "P", "S", "T", "BR", "GR", "PI", "LB", "VI", "GY"
            };
            var wireTypePatterns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "AVSS", "AVS", "FLRY", "T1", "T2", "T3", "T4", "GPT", "CAVS" // Added common types
            };

            // Utility for general noise filtering
            bool IsNoise(string value)
            {
                var v = value.ToUpper();
                // Filter common connector/harness/pin/noise labels
                if (v.StartsWith("PIN") || v.EndsWith(":") || v.Contains("SEAL") || v.Contains("CLIP") ||
                    v.Contains("AMP") || v.Contains("YAZAKI") || v.Contains("KET") || v.Contains("KUM") || v.Contains("SWS") ||
                    v.Contains("NOTE") || v.Contains("SPECIFICATION") || v.Contains("ASSY") || v.Contains("DESCRIPTION"))
                {
                    return true;
                }
                return false;
            }
            // --- End Parsing Definitions ---

            var texts = new List<(string value, double x, double y)>();
            var textEntities = dxf.Entities.All.Where(e => e is netDxf.Entities.Text || e is netDxf.Entities.MText);

            foreach (var ent in textEntities)
            {
                string val = null;
                double x = 0, y = 0;

                if (ent is netDxf.Entities.Text textEntity)
                {
                    val = textEntity.Value;
                    x = textEntity.Position.X;
                    y = textEntity.Position.Y;
                }
                else if (ent is netDxf.Entities.MText mTextEntity)
                {
                    val = mTextEntity.Value;
                    x = mTextEntity.Position.X;
                    y = mTextEntity.Position.Y;
                }

                if (val != null)
                {
                    // Aggressive cleanup: remove DXF formatting codes (\P, \H, \W, etc.) and trim.
                    var safeVal = Regex.Replace(val, @"\\(P|H|W|S|T|L|O|U|A|C|F)[^;]*;?", "").Trim();
                    safeVal = safeVal.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();

                    if (!string.IsNullOrEmpty(safeVal) && !IsNoise(safeVal))
                    {
                        texts.Add((safeVal, x, y));
                    }
                }
            }

            Console.WriteLine($"Collected {texts.Count} cleaned text-like entities.");

            // Simple grouping by Y coordinate (coarse) to create rows
            var groups = new List<List<(string value, double x, double y)>>();
            double yTol = 5.0; // **Decreased Y Tolerance for stricter row grouping (adjust based on your DXF)**

            // Sort descending by Y first for consistent row order (top to bottom)
            foreach (var t in texts.OrderByDescending(t => t.y).ThenBy(t => t.x))
            {
                // Find a group whose average Y is close enough
                var g = groups.FirstOrDefault(gr => gr.Any() && Math.Abs(gr.Average(i => i.y) - t.y) <= yTol);
                if (g == null) groups.Add(new List<(string, double, double)> { t });
                else g.Add(t);
            }

            Console.WriteLine($"Grouped into {groups.Count} rows based on Y coordinate.");

            // Iterate over the grouped rows
            foreach (var g in groups)
            {
                var sortedGroup = g.OrderBy(item => item.x).ToList();
                var row = table.NewRow();

                // Initialize row with empty strings
                foreach (DataColumn column in table.Columns) row[column.ColumnName] = "";

                row["رديف"] = idx.ToString();

                var rowData = new Dictionary<string, string>();
                var usedValues = new HashSet<string>();

                // Stage 1: Absolute Identifiers (Color, Size+Type)
                foreach (var item in sortedGroup)
                {
                    string value = item.value.Trim().ToUpper();
                    if (usedValues.Contains(value)) continue;

                    // 1. رنگ سيم: (بالاترین اولویت، معمولاً کوتاه و در ابتدا)
                    if (table.Columns.Contains("رنگ سيم") && !rowData.ContainsKey("رنگ سيم") &&
                        singleCharColors.Contains(value))
                    {
                        rowData["رنگ سيم"] = value;
                        usedValues.Add(value);
                    }

                    // 2. تشخیص ترکیبی سايزسيم و نوع سيم (مثلاً 0.35AVSS)
                    else if (!rowData.ContainsKey("نوع سيم"))
                    {
                        bool isWireTypeFound = false;
                        foreach (var pattern in wireTypePatterns.OrderByDescending(p => p.Length))
                        {
                            if (value.EndsWith(pattern))
                            {
                                string sizePart = value.Substring(0, value.Length - pattern.Length).Trim();
                                string cleanValue = sizePart.Replace(',', '.').Replace(" ", string.Empty);

                                // Parse Size
                                if (double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double size) &&
                                    size >= 0.1 && size <= 5.0)
                                {
                                    if (table.Columns.Contains("سايزسيم")) rowData["سايزسيم"] = sizePart;
                                    rowData["نوع سيم"] = pattern;
                                    usedValues.Add(value);
                                    isWireTypeFound = true;
                                    break;
                                }
                                // If the value is ONLY the wire type (e.g., AVSS) and no size part
                                else if (string.IsNullOrEmpty(sizePart) && wireTypePatterns.Contains(value))
                                {
                                    rowData["نوع سيم"] = pattern;
                                    usedValues.Add(value);
                                    isWireTypeFound = true;
                                    break;
                                }
                            }
                        }
                    }
                }

                // Stage 2: Length and Single Size (Numbers)
                foreach (var item in sortedGroup)
                {
                    string value = item.value.Trim().ToUpper();
                    if (usedValues.Contains(value)) continue;

                    // Try to parse as integer (potential length)
                    if (int.TryParse(value, out int length))
                    {
                        // 3. طول برش سيم: (عدد صحیح بزرگ - 50 تا 2500)
                        if (table.Columns.Contains("طول برش سيم") && !rowData.ContainsKey("طول برش سيم") &&
                            length >= 50 && length <= 2500)
                        {
                            rowData["طول برش سيم"] = value;
                            usedValues.Add(value);
                        }
                    }
                    // Try to parse as double (potential size - already attempted in Stage 1, but for single instances)
                    else
                    {
                        string cleanValue = value.Replace(',', '.').Replace(" ", string.Empty);
                        // 4. سايزسيم به صورت تکی (0.1 تا 5.0)
                        if (table.Columns.Contains("سايزسيم") && !rowData.ContainsKey("سايزسيم") &&
                            double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double size) &&
                            size >= 0.1 && size <= 5.0 && cleanValue.Length <= 5)
                        {
                            rowData["سايزسيم"] = value;
                            usedValues.Add(value);
                        }
                    }
                }

                // Stage 3: Wire Code and Connector Names (Strings)
                foreach (var item in sortedGroup)
                {
                    string value = item.value.Trim().ToUpper();
                    if (usedValues.Contains(value)) continue;

                    // 5. کد سیم: (متن ترکیبی)
                    // Logic: Must contain letters, must not be noise, length is constrained
                    if (table.Columns.Contains("کدسیم") && !rowData.ContainsKey("کدسیم") &&
                        value.Length >= 3 && value.Length <= 10 && value.Any(char.IsLetter) &&
                        !IsNoise(value) && !singleCharColors.Contains(value) && !wireTypePatterns.Contains(value))
                    {
                        rowData["کدسیم"] = value;
                        usedValues.Add(value);
                    }

                    // 6. ابتدا / 7. انتها (اتصال‌دهنده):
                    // Logic: Must contain letters, length 4-15, not noise, not a wire type/color/code
                    else if (value.Length >= 4 && value.Length <= 15 && value.Any(char.IsLetter) &&
                             !IsNoise(value) && !singleCharColors.Contains(value) && !wireTypePatterns.Contains(value) &&
                             !value.All(c => char.IsDigit(c)))
                    {
                        if (table.Columns.Contains("ابتدا") && !rowData.ContainsKey("ابتدا"))
                        {
                            rowData["ابتدا"] = value;
                            usedValues.Add(value);
                        }
                        else if (table.Columns.Contains("انتها") && !rowData.ContainsKey("انتها"))
                        {
                            rowData["انتها"] = value;
                            usedValues.Add(value);
                        }
                    }
                }

                // Transfer collected data to the DataTable row
                foreach (var kvp in rowData)
                {
                    row[kvp.Key] = kvp.Value;
                }

                table.Rows.Add(row);
                idx++;
            }
        }

        // Save to Excel
        try
        {
            var outputPath = @"C:\tmp\output.xlsx";
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add(table, "Wires");
                wb.Worksheet(1).Columns().AdjustToContents();
                wb.SaveAs(outputPath);
            }
            Console.WriteLine("✅ Excel saved as output.xlsx");
            Console.WriteLine("test");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR saving Excel: {ex.Message}");
        }
    }
}