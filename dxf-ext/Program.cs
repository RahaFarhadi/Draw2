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


        
    }
}
