using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportItems
{
    public partial class Actions
    {
        private void Actions_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ExportPDF(object sender, RibbonControlEventArgs e) =>
            Models.Excel.ExportPDF();

        private void ExportXLSX(object sender, RibbonControlEventArgs e) =>
            Models.Excel.ExportXLSX();

        private void SendEmail(object sender, RibbonControlEventArgs e)
        {
            Models.Outlook.SendExcel();
        }
    }
}
