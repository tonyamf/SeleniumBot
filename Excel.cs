using Microsoft.Office.Interop.Excel;
using System;
using _Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumBot
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(String path, int Sheet)
        {
            this.path = path;
            excel.Visible = true;

            wb = excel.Workbooks.Open(path, 0, false, 5, "", "", false,
                XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ws = wb.Worksheets[Sheet];
        }
        public void Write(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s;
        }
        public void Writeint(int i, int j, double s)
        {
            ws.Cells[i, j].Value2 = s;
        }

        public void save()
        {
            wb.Save();
        }

    }
}
