using Microsoft.Office.Interop.Excel;
using System;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Application = Microsoft.Office.Interop.Outlook.Application;
using MailItem = Microsoft.Office.Interop.Outlook.MailItem;
using Microsoft.Office.Interop.Outlook;

namespace ExportItems.Models
{
    public class Outlook
    {
        public static void SendExcel()
        {
            //instancia do Outlook
            Application outlookApp = new Application();
            MailItem email = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

            //instacia do Excel
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            //destino do email e assunto.
            email.To = currentSheet.Cells[5, 1].Value;
            email.CC = currentSheet.Cells[6, 1].Value;
            email.Subject = currentSheet.Cells[8, 1].Value;



            //conteudo do email
            email.Body = "Hello " + currentSheet.Cells[7, 1].value + ", \n"
                + "Please find attached the updated release.";

            string pdf = Excel.ExportPDF();
            string excel = Excel.ExportXLSX();

            email.Attachments.Add(excel);
            email.Attachments.Add(pdf);
            email.Display(true);


            // Libere os recursos
            System.Runtime.InteropServices.Marshal.ReleaseComObject(email);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(currentSheet);


        }
    }
}
