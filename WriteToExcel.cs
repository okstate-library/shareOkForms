using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace shareOkForms
{
    class WriteToExcel
    {
        public bool writeExcel(System.Data.DataTable dt, String outputFileName,String outputDirectory)
        {
            if (!System.IO.File.Exists(outputDirectory))
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }
            FileInfo f = new FileInfo(outputDirectory + outputFileName);
            using (ExcelPackage pck = new ExcelPackage(f))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Accounts");
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                pck.Save();
            }
            return false;
        }
    }
}
