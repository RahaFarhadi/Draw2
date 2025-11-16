using System;
using System.Linq;
using netDxf;
using netDxf.Entities;
using netDxf.Tables;
using ClosedXML.Excel;

class Program
{
    static void Main(string[] args)
    {
        string dxfPath = @"C:\tmp\test2-a.dxf";
        string excelPath = @"C:\tmp\output3.xlsx";

        // بارگذاری فایل DXF
        DxfDocument dxf = DxfDocument.Load(dxfPath);
        if (dxf == null)
        {
            Console.WriteLine("فایل DXF بارگذاری نشد.");
            return;
        }

        // ایجاد فایل Excel
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Wires");

        // هدرها
        worksheet.Cell(1, 1).Value = "ردیف";
        worksheet.Cell(1, 2).Value = "کد سیم";
        worksheet.Cell(1, 3).Value = "رنگ سیم";
        worksheet.Cell(1, 4).Value = "نوع سیم";
        worksheet.Cell(1, 5).Value = "اندازه سیم";

        int row = 2;
        int index = 1;

        // پیمایش Line ها
        foreach (Line line in dxf.Lines)
        {
            string wireCode = line.Layer?.Name ?? "Unknown";
            string wireColor = line.Color.IsByLayer ? line.Layer.Color.ToString() : line.Color.ToString();
            string wireType = line.Linetype?.Name ?? "Default";
            double wireLength = line.StartPoint.DistanceTo(line.EndPoint);

            worksheet.Cell(row, 1).Value = index++;
            worksheet.Cell(row, 2).Value = wireCode;
            worksheet.Cell(row, 3).Value = wireColor;
            worksheet.Cell(row, 4).Value = wireType;
            worksheet.Cell(row, 5).Value = wireLength;
            row++;
        }

        // پیمایش Polyline ها
        foreach (LwPolyline poly in dxf.LwPolylines)
        {
            string wireCode = poly.Layer?.Name ?? "Unknown";
            string wireColor = poly.Color.IsByLayer ? poly.Layer.Color.ToString() : poly.Color.ToString();
            string wireType = poly.Linetype?.Name ?? "Default";
            double wireLength = poly.Length();

            worksheet.Cell(row, 1).Value = index++;
            worksheet.Cell(row, 2).Value = wireCode;
            worksheet.Cell(row, 3).Value = wireColor;
            worksheet.Cell(row, 4).Value = wireType;
            worksheet.Cell(row, 5).Value = wireLength;
            row++;
        }

        // پیمایش BlockReference ها (اگر سیم‌ها در بلاک باشند)
        foreach (BlockReference block in dxf.Blocks.SelectMany(b => b.Entities.OfType<BlockReference>()))
        {
            string wireCode = block.Layer?.Name ?? "Unknown";
            string wireColor = block.Color.IsByLayer ? block.Layer.Color.ToString() : block.Color.ToString();
            string wireType = block.Linetype?.Name ?? "Default";
            double wireSize = block.Scale.X; // یا هر ویژگی مرتبط

            worksheet.Cell(row, 1).Value = index++;
            worksheet.Cell(row, 2).Value = wireCode;
            worksheet.Cell(row, 3).Value = wireColor;
            worksheet.Cell(row, 4).Value = wireType;
            worksheet.Cell(row, 5).Value = wireSize;
            row++;
        }

        // ذخیره فایل Excel
        workbook.SaveAs(excelPath);
        Console.WriteLine("اطلاعات سیم‌ها در فایل Excel ذخیره شد: " + excelPath);
    }
}
