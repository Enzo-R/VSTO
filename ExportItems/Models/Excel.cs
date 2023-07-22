using ExportItems;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportItems.Models
{
    public class Excel
    {
        public static string ExportPDF()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            DateTime dateTime = DateTime.Today;
            string date = dateTime.ToString("d");
            string dateValidate = date.Replace("/", ".");
            string name = currentSheet.Name;
            string filename = name + " - " + dateValidate+".pdf";
            string path = @"S:\Log_Planej_Adm\PERSONAL\Matheus Rodrigues\1- Releases\" + filename;

            if (!File.Exists(path))
            {
                currentSheet.ExportAsFixedFormat
                (XlFixedFormatType.xlTypePDF, path, XlFixedFormatQuality.xlQualityMinimum);
                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(currentSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.getActiveApp());
                
            }

            return path;
        }

        public static string ExportXLSX()
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            DateTime dateTime = DateTime.Today;
            string date = dateTime.ToString("d");
            string dateValidate = date.Replace("/", ".");
            string name = currentSheet.Name;
            string filename = name + " - " + dateValidate + ".xlsx";
            string path = @"S:\Log_Planej_Adm\PERSONAL\Matheus Rodrigues\0- Releases excel\" + filename;

            if (!File.Exists(path))
            {
                // Cria um novo workbook
                Application excelApp = Globals.ThisAddIn.getActiveApp();
                Workbook newWorkbook = excelApp.Workbooks.Add();
                currentSheet.Copy(Type.Missing, newWorkbook.Sheets[1]);
                newWorkbook.Sheets["Planilha1"].Delete();
                newWorkbook.SaveAs(path);
                newWorkbook.Close(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(currentSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.getActiveApp());
            }

            return path;
        }
    }
}
