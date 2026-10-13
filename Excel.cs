using ClosedXML.Excel;
using System;

namespace SeleniumBot
{
    // Writes cells to an existing .xlsx file (was Microsoft.Office.Interop.Excel, which needs Windows + Office).
    class Excel
    {
        string path = "";
        XLWorkbook wb;
        IXLWorksheet ws;
        public Excel(String path, int Sheet)
        {
            this.path = path;
            wb = new XLWorkbook(path);
            ws = wb.Worksheet(Sheet);
        }
        public void Write(int i, int j, string s)
        {
            ws.Cell(i, j).Value = s;
        }
        public void Writeint(int i, int j, double s)
        {
            ws.Cell(i, j).Value = s;
        }

        public void save()
        {
            wb.Save();
        }

    }
}
