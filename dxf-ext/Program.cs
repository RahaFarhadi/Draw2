//using System;

//using System.IO;

//using System.Linq;

//using System.Data;

//using System.Text.Json;

//using System.Collections;

//using System.Collections.Generic;

//using System.Reflection;

//using netDxf;

//using ClosedXML.Excel;

//using System.Globalization;



////class MappingColumn { public string col { get; set; } public string attr { get; set; } public string source { get; set; } }

////class Mapping { public List<MappingColumn> columns { get; set; } }



//class Program

//{

//    static void Main(string[] args)

//    {

//        Console.OutputEncoding = System.Text.Encoding.UTF8;

//        string dxfPath = "C:\\tmp\\sample-dxf.dxf";

//        string mappingPath = "C:\\Users\\Digibod.ir\\source\\repos\\dxf-ext\\dxf-ext\\mapping.json";



//        if (!File.Exists(dxfPath)) { Console.WriteLine($"ERROR: DXF file not found: {dxfPath}"); return; }

//        if (!File.Exists(mappingPath)) { Console.WriteLine($"ERROR: mapping.json not found: {mappingPath}"); return; }



//        Mapping mapping = JsonSerializer.Deserialize<Mapping>(File.ReadAllText(mappingPath));



//        // Load DXF

//        DxfDocument dxf = null;

//        try

//        {

//            dxf = DxfDocument.Load(dxfPath);

//            Console.WriteLine($"✅ DXF loaded. Entities count: {dxf.Entities.All.Count()}");

//        }

//        catch (Exception ex)

//        {

//            Console.WriteLine($"ERROR loading DXF: {ex.Message}");

//            return;

//        }



//        // Utility: get enumerable from a possible property or method using reflection

//        IEnumerable SafeGetEnumerable(object owner, string memberName)

//        {

//            if (owner == null) return Enumerable.Empty<object>();



//            var t = owner.GetType();



//            // Try property

//            var prop = t.GetProperty(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);

//            if (prop != null)

//            {

//                var val = prop.GetValue(owner);

//                if (val is IEnumerable e) return e.Cast<object>();

//            }



//            // Try method (no parameters)

//            var m = t.GetMethod(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase, null, Type.EmptyTypes, null);

//            if (m != null)

//            {

//                var val = m.Invoke(owner, null);

//                if (val is IEnumerable e2) return e2.Cast<object>();

//            }



//            // Try field (rare)

//            var field = t.GetField(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);

//            if (field != null)

//            {

//                var val = field.GetValue(owner);

//                if (val is IEnumerable e3) return e3.Cast<object>();

//            }



//            return Enumerable.Empty<object>();

//        }



//        // Debug: show counts / types for Entities and Inserts

//        var entitiesEnum = SafeGetEnumerable(dxf, "Entities") ?? Enumerable.Empty<object>();

//        var insertsEnum = SafeGetEnumerable(dxf, "Inserts") ?? Enumerable.Empty<object>();



//        int entitiesCount = entitiesEnum.Cast<object>().Count();

//        int insertsCount = insertsEnum.Cast<object>().Count();



//        Console.WriteLine($"DXF loaded. Entities count (if available): {entitiesCount}");

//        Console.WriteLine($"DXF loaded. Inserts count (if available): {insertsCount}");





//        var table = new DataTable("Wires");

//        foreach (var mc in mapping.columns) table.Columns.Add(mc.col);



//        int idx = 1;

//        bool dataExtractedFromInserts = false; // پرچم جدید



//        // استفاده از موجودیت‌های کلی برای فیلتر کردن Inserts (ایمن‌ترین روش)

//        var inserts = dxf.Entities.All.OfType<netDxf.Entities.Insert>().ToList();



//        if (inserts.Count > 0)

//        {

//            Console.WriteLine($"✅ Found {inserts.Count} Insert entities. Attempting to read Attributes.");



//            // Iterate inserts (block references) if any

//            foreach (var insObj in inserts)

//            {

//                // Attributeها را فیلتر کنید

//                //var attributes = insObj.Attributes.OfType<netDxf.Entities.AttributeDefinition>().ToList();

//                var attributes = insObj.Attributes.ToList();

//                Console.WriteLine($"--- Attributes for Insert at ({insObj.Position.X:F0}, {insObj.Position.Y:F0}) ---");

//                foreach (var a in attributes)

//                {

//                    Console.WriteLine($"Tag: {a.Tag}, Value: {a.Value}");

//                }

//                Console.WriteLine($"----------------------------------");



//                //if (attributes.Count > 0)

//                //{

//                //    dataExtractedFromInserts = true;

//                //    var row = table.NewRow();



//                //    foreach (var mc in mapping.columns)

//                //    {

//                //        if (mc.source == "auto_index")

//                //        {

//                //            row[mc.col] = idx;

//                //            continue;

//                //        }

//                //        if (!string.IsNullOrEmpty(mc.attr))

//                //        {

//                //            var attribute = attributes.FirstOrDefault(a =>

//                //                string.Equals(a.Tag, mc.attr, StringComparison.OrdinalIgnoreCase)

//                //            );



//                //            if (attribute != null)

//                //            {

//                //                row[mc.col] = attribute.Value;

//                //            }

//                //            else

//                //            {

//                //                row[mc.col] = "";

//                //            }

//                //        }

//                //    }

//                //    table.Rows.Add(row);

//                //    idx++;

//                //}

//                // اگر Attribute وجود نداشت، این Insert نادیده گرفته می‌شود.

//            }

//        }





//        // منطق Fallback: اگر داده‌ای از Insertها استخراج نشد (یا اگر اصلا Insertی نبود)

//        if (!dataExtractedFromInserts)

//        {

//            Console.WriteLine("⚠️ No structured attribute data found. Falling back to parsing Text/MText entities.");



//            var singleCharColors = new HashSet<string>(StringComparer.OrdinalIgnoreCase)

//            {

//                "R", "B", "Y", "L", "W", "G", "V", "O", "P", "S", "T", "BR", "GR", "PI", "LB", "VI", "GY"

//            };

//            var wireTypePatterns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)

//            {

//                "AVSS", "AVS", "FLRY", "T1", "T2", "T3", "T4"

//            };

//            bool IsPinNumberOrNoise(string value)

//            {

//                var v = value.ToUpper();



//                // فیلتر شماره پین‌ها و خطوط جداکننده

//                if (v.StartsWith("PIN") || v.EndsWith(":") || v.Contains("SEAL") || v.Contains("------------"))

//                {

//                    return true;

//                }



//                // فیلتر نویزهای عمومی و نام‌های سازنده

//                if (v.Contains("AMP") || v.Contains("YAZAKI") || v.Contains("KET") || v.Contains("KUM") || v.Contains("SWS") ||

//                    v.Contains("CLIP") || v.Contains("INTER CONNECTION") || v.Contains("NOTE") ||

//                    v.Contains("SWITCH") || v.Contains("DOOR LOCK") || v.Contains("SPEAKER") ||

//                    v.Contains("ELECTRIC MIRROR") || v.Contains("BACK VIEW") || v.Contains("ASSY") || v.Contains("SPECIFICATION"))

//                {

//                    return true;

//                }



//                // فیلتر طول برش (اعداد صحیح بین 50 تا 2500)

//                if (int.TryParse(v, out int length) && length >= 50 && length <= 2500)

//                {

//                    return true;

//                }



//                // فیلتر رنگ‌های تک یا دو حرفی (اگرچه در بالا هم فیلتر شده‌اند، اینجا برای ایمنی بیشتر است)

//                if ((v.Length >= 1 && v.Length <= 2) && singleCharColors.Contains(v))

//                {

//                    return true;

//                }



//                return false;

//            }



//            var texts = new List<(string value, double x, double y)>();



//            var textEntities = dxf.Entities.All.Where(e => e is netDxf.Entities.Text || e is netDxf.Entities.MText);



//            foreach (var ent in textEntities)

//            {

//                string val = null;

//                double x = 0, y = 0;



//                if (ent is netDxf.Entities.Text textEntity)

//                {

//                    val = textEntity.Value;

//                    x = textEntity.Position.X;

//                    y = textEntity.Position.Y;

//                }

//                else if (ent is netDxf.Entities.MText mTextEntity)

//                {

//                    val = mTextEntity.Value;

//                    x = mTextEntity.Position.X;

//                    y = mTextEntity.Position.Y;

//                }



//                if (val != null)

//                {

//                    var safeVal = (val ?? string.Empty).Trim().Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Replace("\\P", " ").Trim();

//                    if (!string.IsNullOrEmpty(safeVal))

//                        texts.Add((safeVal, x, y));

//                }

//            }



//            Console.WriteLine($"Collected {texts.Count} text-like entities.");



//            // Simple grouping by Y coordinate (coarse) to create rows

//            var groups = new List<List<(string value, double x, double y)>>();

//            double yTol = 12.0; // Tolerance 10.0 is used for grouping rows



//            foreach (var t in texts.OrderByDescending(t => t.y))

//            {

//                var g = groups.FirstOrDefault(gr => gr.Any() && Math.Abs(gr[0].y - t.y) <= yTol);

//                if (g == null) groups.Add(new List<(string, double, double)> { t });

//                else g.Add(t);

//            }



//            Console.WriteLine($"Grouped into {groups.Count} rows based on Y coordinate.");



//            // Flags for sequential assignment of Start/End

//            bool codeAssigned;

//            bool startAssigned;



//            foreach (var g in groups)

//            {

//                var sortedGroup = g.OrderBy(item => item.x).ToList();

//                var row = table.NewRow();

//                codeAssigned = false;

//                startAssigned = false;



//                Console.WriteLine($"DEBUG ROW {idx}: {string.Join(" | ", sortedGroup.Select(item => $"'{item.value}'"))}");



//                row["رديف"] = idx.ToString();

//                if (table.Columns.Contains("کدسیم")) row["کدسیم"] = "";

//                if (table.Columns.Contains("رنگ سيم")) row["رنگ سيم"] = "";

//                if (table.Columns.Contains("سايزسيم")) row["سايزسيم"] = "";

//                if (table.Columns.Contains("نوع سيم")) row["نوع سيم"] = "";

//                if (table.Columns.Contains("طول برش سيم")) row["طول برش سيم"] = "";

//                if (table.Columns.Contains("ابتدا")) row["ابتدا"] = "";

//                if (table.Columns.Contains("انتها")) row["انتها"] = "";



//                // ** منطق تخصیص اولویت‌بندی شده و سخت‌گیرانه **

//                var usedValues = new HashSet<string>();

//                foreach (var item in sortedGroup)

//                {

//                    string value = item.value.Trim().ToUpper();

//                    // **پاکسازی تهاجمی برای حذف کدهای قالب‌بندی DXF**

//                    value = value.Replace(@"\PXQC;", "").Trim();

//                    if (usedValues.Contains(value)) continue;





//                    // 1. رنگ سيم: (بالاترین اولویت)

//                    if (table.Columns.Contains("رنگ سيم") && string.IsNullOrEmpty(row["رنگ سيم"].ToString()) &&

//                      singleCharColors.Contains(value))

//                    {

//                        row["رنگ سيم"] = value;

//                    }



//                    // 2. نوع سيم (گرید حرارتی): (اولویت دوم)

//                    //else if (table.Columns.Contains("نوع سيم") && string.IsNullOrEmpty(row["نوع سيم"].ToString()))

//                    //{

//                    //    bool isWireTypeFound = false;

//                    //    // بررسی می‌کنیم که آیا مقدار با یکی از الگوهای نوع سیم خاتمه می‌یابد (الگوهای طولانی‌تر اولویت دارند).

//                    //    foreach (var pattern in wireTypePatterns.OrderByDescending(p => p.Length))

//                    //    {

//                    //        if (value.EndsWith(pattern))

//                    //        {

//                    //            row["نوع سيم"] = pattern;

//                    //            isWireTypeFound = true;

//                    //            break;

//                    //        }

//                    //    }



//                    //    // در صورتی که خود مقدار، یک نوع سیم کامل (مانند AVSS) باشد.

//                    //    if (!isWireTypeFound && wireTypePatterns.Contains(value))

//                    //    {

//                    //        row["نوع سيم"] = value;

//                    //    }

//                    //}

//                    // 2. تشخیص ترکیبی سايزسيم و نوع سيم (مثلاً 0.35AVSS)

//                    else if (table.Columns.Contains("نوع سيم") && string.IsNullOrEmpty(row["نوع سيم"].ToString()))

//                    {

//                        bool isWireTypeFound = false;

//                        foreach (var pattern in wireTypePatterns.OrderByDescending(p => p.Length))

//                        {

//                            if (value.EndsWith(pattern))

//                            {

//                                // جدا کردن سايز

//                                string sizePart = value.Substring(0, value.Length - pattern.Length).Trim();



//                                string cleanValue = sizePart.Replace(',', '.').Replace(" ", string.Empty);

//                                if (double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double size) &&

//                                    size >= 0.1 && size <= 5.0)

//                                {

//                                    if (table.Columns.Contains("سايزسيم")) row["سايزسيم"] = sizePart;

//                                    row["نوع سيم"] = pattern;

//                                    usedValues.Add(value);

//                                    isWireTypeFound = true;

//                                    break;

//                                }

//                                // اگر قسمت سایز یک عدد معتبر نبود، فقط نوع سیم را ذخیره می‌کنیم

//                                else if (string.IsNullOrEmpty(sizePart))

//                                {

//                                    row["نوع سيم"] = pattern;

//                                    usedValues.Add(value); // فقط نوع سیم

//                                    isWireTypeFound = true;

//                                    break;

//                                }

//                            }

//                        }

//                        // در صورتی که خود مقدار، یک نوع سیم کامل (مانند AVSS) باشد.

//                        if (!isWireTypeFound && wireTypePatterns.Contains(value))

//                        {

//                            row["نوع سيم"] = value;

//                            usedValues.Add(value);

//                        }

//                    }

//                    // 3. تشخیص سايزسيم به صورت تکی (اولویت بالا برای استخراج اعداد)

//                    else if (table.Columns.Contains("سايزسيم") && string.IsNullOrEmpty(row["سايزسيم"].ToString()))

//                    {

//                        string cleanValue = value.Replace(',', '.').Replace(" ", string.Empty);

//                        // کاملاً عدد باشد و رنج منطقی سایز سیم (0.1 تا 5.0)

//                        if (cleanValue.All(c => char.IsDigit(c) || c == '.'))

//                        {

//                            if (double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double size) &&

//                                size >= 0.1 && size <= 5.0 && cleanValue.Length <= 5)

//                            {

//                                row["سايزسيم"] = value;

//                                usedValues.Add(value);

//                            }

//                        }

//                    }

//                    // 3. طول برش سيم: (عدد صحیح بزرگ) (اولویت سوم)





//                    // 4. کد سیم: (متن ترکیبی - منطق **اصلاح شده**) (اولویت چهارم)

//                    else if (table.Columns.Contains("کدسیم") && string.IsNullOrEmpty(row["کدسیم"].ToString()) &&

//                             value.Length >= 3 && value.Length <= 8 &&

//                             value.Any(char.IsLetter) && !value.Contains(".") && !value.Contains("-") && value.All(c => char.IsLetterOrDigit(c) || c == '.' || c == '-'))

//                    {

//                        row["کدسیم"] = value;

//                    }

//                    //else if (table.Columns.Contains("کدسیم") && string.IsNullOrEmpty(row["کدسیم"].ToString()))

//                    //{

//                    //    // فیلترها:

//                    //    bool isTooShort = value.Length < 3;

//                    //    bool isJustNumbersOrSize = value.All(c => char.IsDigit(c) || c == '.' || c == ',');



//                    //    // شرط اصلی: حداقل 3 کاراکتر، کاملاً عدد یا سایز نباشد، نویز نباشد و حداقل یک حرف یا نقطه داشته باشد (برای GPT.1)

//                    //    if (!isTooShort && !isJustNumbersOrSize && !IsPinNumberOrNoise(value))

//                    //    {

//                    //        // اطمینان از اینکه واقعاً یک کد سیم است و نه فقط یک کلمه طولانی

//                    //        if (value.Any(char.IsLetter) || value.Contains('.'))

//                    //        {

//                    //            row["کدسیم"] = value;

//                    //            usedValues.Add(value);

//                    //        }

//                    //    }

//                    //}





//                    //else if (table.Columns.Contains("کدسیم") && string.IsNullOrEmpty(row["کدسیم"].ToString()) &&

//                    // value.Length >= 3 && value.Length <= 10 &&

//                    // value.Any(char.IsLetter) &&

//                    // !value.Contains("PIN") && !value.Contains("SEAL") && !value.Contains("NOTE") &&

//                    // !value.Contains("BUSS") && !value.Contains("CLIP") &&

//                    // !value.Contains("AMP") && !value.Contains("YAZAKI") && !value.Contains("KET") && !value.Contains("KUM") && !value.Contains("SWS") && // Exclude connector/clip manufacturer names

//                    // !value.All(char.IsDigit)) // Not purely a number

//                    //{

//                    //    row["کدسیم"] = value;

//                    //    codeAssigned = true;

//                    //}



//                    // 5. ابتدا (شماره سوکت 1): (بعد از تخصیص کد سیم)

//                    // فرض می‌کنیم اولین متن متنی یونیک بعد از استخراج همه فیلدهای دیگر، "ابتدا" است.



//                    else if (table.Columns.Contains("ابتدا") && string.IsNullOrEmpty(row["ابتدا"].ToString()) && codeAssigned &&

//                             value.Length >= 4 && value.Length <= 15 && value.Any(char.IsLetter) &&

//                             !value.Contains("PIN") && !value.Contains("-") && !wireTypePatterns.Contains(value))

//                    {

//                        row["ابتدا"] = value;

//                        startAssigned = true;

//                    }



//                    // 6. انتها (شماره سوکت 2): (بعد از تخصیص ابتدا)

//                    // دومین متن متنی یونیک بعد از "ابتدا" است.

//                    else if (table.Columns.Contains("انتها") && string.IsNullOrEmpty(row["انتها"].ToString()) && startAssigned &&

//                             value.Length >= 4 && value.Length <= 15 && value.Any(char.IsLetter) &&

//                             !value.Contains("PIN") && !value.Contains("-") && !wireTypePatterns.Contains(value))

//                    {

//                        row["انتها"] = value;

//                    }

//                    if (table.Columns.Contains("طول برش سيم") && string.IsNullOrEmpty(row["طول برش سيم"].ToString()))

//                    {

//                        if (int.TryParse(value, out int length))

//                        {

//                            // شرط طول برش: معمولاً بین 50 تا 2500 میلی‌متر است.

//                            if (length >= 50 && length <= 2500)

//                            {

//                                row["طول برش سيم"] = value;

//                            }

//                        }

//                    }



//                    // نادیده گرفتن سایر متون

//                }



//                table.Rows.Add(row);

//                idx++;

//            }

//        }

//        // Save to Excel

//        try

//        {

//            var outputPath = @"C:\tmp\output.xlsx";

//            using (var wb = new XLWorkbook())

//            {

//                wb.Worksheets.Add(table, "Wires");

//                wb.Worksheet(1).Columns().AdjustToContents();

//                wb.SaveAs(outputPath);

//            }

//            Console.WriteLine("✅ Excel saved as output.xlsx");

//        }

//        catch (Exception ex)

//        {

//            Console.WriteLine($"ERROR saving Excel: {ex.Message}");

//        }

//    }

//}