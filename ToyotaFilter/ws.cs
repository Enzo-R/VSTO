using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ToyotaFilter
{
    public class ws
    {

        static string[] filterCriteria = new string[] 
                {
                "1027659",
                "1027946",
                "1027995",
                "1028567",
                "1033613345",
                "1033614345",
                "1059659",
                "1070960JN7",
                "1085204",
                "1087494483",
                "1087495483",
                "1113370",
                "1119087",
                "1123477483",
                "1124938",
                "1126576483",
                "1133077",
                "1137111JN7",
                "1140742",
                "1144807",
                "1147650",
                "2517433",
                "305313710",
                "4012409",
                "7ADC7176",
                "7ADC9002",
                "7ADC9003",
                "7ADC9016",
                "7ADG2286Z03",
                "7LB05067"
                };

        public static int GetRange(string range, Worksheet sheet, int count =0)
        {
            //get the sheet range
            Range cellSelect = sheet.Range[range];
            Range sl = sheet.Range[cellSelect, cellSelect.End[XlDirection.xlDown]];
            int rows = sl.SpecialCells(XlCellType.xlCellTypeVisible).Count + count;

            return rows;
        }


        public static void OpenGIT(string path, string sheet)
        {
            OpenShortageReport();

            //Current file
            Worksheet ws = Globals.ThisAddIn.getActiveWorkbook().Sheets["TOYOTA SHORTAGE REPORT"];
            ws.Activate();

            //Open file
            Application excelApp = Globals.ThisAddIn.getActiveApp();
            Workbook tempWb = excelApp.Workbooks.Open(path, false);
            Worksheet temp = Globals.ThisAddIn.getActiveWorkbook().Sheets.Add();
            temp.Name = "temp";
            Worksheet wsTemp = tempWb.Sheets[sheet];
            temp.Activate();

            //Manipulating objects
            int rows0 = GetRange("A3", wsTemp, 3);
            Range o3 = ws.Range["O3:O" + rows0]; o3.NumberFormat = "dd/mmm";
            Range p3 = ws.Range["P3:P" + rows0]; p3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            Range q3 = ws.Range["Q3:Q" + rows0]; q3.NumberFormat = "dd/mmm";
            Range r3 = ws.Range["R3:R" + rows0]; r3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            Range s3 = ws.Range["S3:S" + rows0]; s3.NumberFormat = "dd/mmm";
            Range t3 = ws.Range["T3:T" + rows0]; t3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            Range u3 = ws.Range["U3:U" + rows0]; u3.NumberFormat = "dd/mmm";
            Range v3 = ws.Range["V3:V" + rows0]; v3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            Range w3 = ws.Range["W3:W" + rows0]; w3.NumberFormat = "dd/mmm";
            Range x3 = ws.Range["X3:X" + rows0]; x3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";

            //filter PNs wsTemp
            int rows1 = GetRange("I2", wsTemp, 1);
            Range i2 = wsTemp.Range["I2:I" + rows1];
            Range l2 = wsTemp.Range["L2:L" + rows1];
            Range p2 = wsTemp.Range["P2:P" + rows1];

            for (int row = 0; row < filterCriteria.Length; row++)
            {
                string valorAtual = filterCriteria[row];
                Range found = i2.Find(valorAtual);
                if (found != null)
                {
                    if (i2.AutoFilter(9, valorAtual))
                    {
                        //wsTemp.Cells.SpecialCells(XlCellType.xlCellTypeVisible).Value;
                        wsTemp.Cells.SpecialCells(XlCellType.xlCellTypeVisible).SpecialCells(XlCellType.xlCellTypeConstants).Copy();

                        if (temp.Cells[2].value != null)
                        {
                            Range lastRow = temp.Cells.Find("*", System.Reflection.Missing.Value,
                                                                XlFindLookIn.xlValues, XlLookAt.xlPart,
                                                                XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                                                                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            int lastRowNumber = lastRow.Row + 1;

                            //colar
                            Range end2 = temp.Range["A" + lastRowNumber];
                            end2.PasteSpecial(XlPasteType.xlPasteAll);
                            Clipboard.Clear();
                        }
                        else
                        {
                            Range f = temp.Range["A1"];
                            f.PasteSpecial(XlPasteType.xlPasteAll);
                            Clipboard.Clear();
                        }
                    }
                    wsTemp.ShowAllData();
                }
            }
            temp.Columns.EntireColumn.AutoFit();

            int[] b = { 1, 0, 2, 4, 6, 8 };

            //filter PNs
            int rows = GetRange("I2", temp, 1);
            Range i22 = temp.Range["I2:I" + rows];
            Range k22 = temp.Range["K2:K" + rows];
            Range l22 = temp.Range["L2:L" + rows];
            Range p22 = temp.Range["P2:P" + rows];
            Range Frow = temp.Rows[1];
            Frow.EntireRow.Select();
            Frow.AutoFilter(1);

            for (int col = 1; col < b.Length; col++)
            {
                int c = b[col];

                for (int row = 0; row < filterCriteria.Length; row++)
                {
                    string valorAtual = filterCriteria[row];
                    Range found = i22.Find(valorAtual);

                    if (found != null)
                    {
                        if (i22.AutoFilter(9, valorAtual))
                        {

                            //pegar o primeiro valor QTD
                            Range visibleCellsQTD = l22.SpecialCells(XlCellType.xlCellTypeVisible);
                            Range firstCellQTD = visibleCellsQTD.Cells[col];
                            if (firstCellQTD.EntireRow.Hidden != true)
                            {
                                //p22.SpecialCells(XlCellType.xlCellTypeVisible).Sort(p22, XlSortOrder.xlAscending);
                                Range QTD = ws.Cells[3 + row, 16 + c];
                                QTD.Value = firstCellQTD.Value;
                            }

                            //pegar o primeiro valor ETA
                            Range visibleCellsETA = p22.SpecialCells(XlCellType.xlCellTypeVisible);
                            Range firstCellETA = visibleCellsETA.Cells[col];
                            if (firstCellETA.EntireRow.Hidden != true)
                            {
                                //p22.SpecialCells(XlCellType.xlCellTypeVisible).Sort(p22, XlSortOrder.xlAscending);
                                Range ETA = ws.Cells[3 + row, 15 + c];
                                ETA.Value = firstCellETA.Value;
                            }
                        }
                        temp.ShowAllData();
                    }
                }
            }
            ws.Columns["O:X"].EntireColumn.AutoFit();
            ws.Activate();

            temp.Delete();

            ReleaseObject(excelApp);
            ReleaseObject(ws);
            ReleaseObject(tempWb);
            ReleaseObject(wsTemp);
        }



        public static void OpenShortageReport()
        {
            //temp
            Worksheet temp = Globals.ThisAddIn.getActiveWorkbook().Sheets.Add();
            temp.Name = "temp";
            //toy
            Worksheet toyota = Globals.ThisAddIn.getActiveWorkbook().Sheets.Add();
            toyota.Name = "TOYOTA SHORTAGE REPORT";
            //report
            Worksheet report = Globals.ThisAddIn.getActiveWorkbook().Sheets["SHORTAGE REPORT"];
            report.Activate();

            //passar os dados para nova planilha
            int a = GetRange("A2", report);
            report.Range["A2:J" + a].Copy();
            temp.Cells[1].PasteSpecial(XlPasteType.xlPasteValues);
            temp.Activate();

            //temp Ranges ok
            Range Frow = temp.Rows[1];
            Frow.EntireRow.Select();
            Frow.AutoFilter(1);

            int b = GetRange("A2", temp,1);
            Range a2 = temp.Range["A2:A"+b];
            Range firstR = temp.Range["A2:G"+b];
            Range secondR = temp.Range["I2:J"+b];

            a2.AutoFilter(1, filterCriteria, XlAutoFilterOperator.xlFilterValues);

            firstR.SpecialCells(XlCellType.xlCellTypeVisible).Copy();
            toyota.Cells[3,1].PasteSpecial(XlPasteType.xlPasteValues);

            secondR.SpecialCells(XlCellType.xlCellTypeVisible).Copy();
            toyota.Range["K3:L32"].PasteSpecial(XlPasteType.xlPasteValues);

            //toyota ranges
            toyota.Activate();
            int c = GetRange("A2", temp, 2);
            Range f3 = toyota.Range["F3:F" + c];
            Range g3 = toyota.Range["G3:G" + c];
            Range h3 = toyota.Range["H3:H" + c];
            Range i3 = toyota.Range["I3:I" + c];
            Range j3 = toyota.Range["J3:J" + c];
            Range k3 = toyota.Range["K3:K" + c];
            Range l3 = toyota.Range["L3:L" + c];
            Range m3 = toyota.Range["M3:M" + c];
            Range n3 = toyota.Range["N3:N" + c];
            Range I2 = toyota.Range["I2"];

            //FORMULAS
            h3.Formula = @"=F3/G3";
            i3.Formula = @"=(H3+$i$2)+2";
            j3.Formula = @"=K3/G3";
            m3.Formula = @"=IF(P3<>0,P3+R3+T3+V3+X3,0)";
            n3.Formula = @"=M3/G3";
            I2.Formula = @"=TODAY()";

            #region columns edit

            toyota.Range["A1"].Value = "PN";
            toyota.Range["B1"].Value = "DESCRIÇÃO";
            toyota.Range["C1"].Value = "ORIGEM";
            toyota.Range["D1"].Value = "CLIENTE";
            toyota.Range["E1"].Value = "PLANNER";
            toyota.Range["F1"].Value = "ESTOQUE ATUAL";
            toyota.Range["G1"].Value = "USO DIA";
            toyota.Range["H1"].Value = "DOH ATUAL";
            toyota.Range["I1"].Value = "COBERTURA";
            toyota.Range["J1"].Value = "DOH TARGET";
            toyota.Range["K1"].Value = "EST. MINÍMO";
            toyota.Range["L1"].Value = "EST. MÁXIMO";
            toyota.Range["M1"].Value = "TRANSITO";
            toyota.Range["N1"].Value = "DIAS EST.TRANSITO";
            toyota.Range["O1"].Value = "ETA 1";
            toyota.Range["P1"].Value = "QTD";
            toyota.Range["Q1"].Value = "ETA 2";
            toyota.Range["R1"].Value = "QTD";
            toyota.Range["S1"].Value = "ETA 3";
            toyota.Range["T1"].Value = "QTD";
            toyota.Range["U1"].Value = "ETA 4";
            toyota.Range["V1"].Value = "QTD";
            toyota.Range["W1"].Value = "ETA 5";
            toyota.Range["X1"].Value = "QTD";

            f3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            g3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            h3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            i3.NumberFormat = @"m/d/yyyy";
            j3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            k3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            l3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            m3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";
            n3.NumberFormat = @"_-* #,##0_-;-* #,##0_-;_-* " + "-" + "??_-;_-@_-";

            toyota.Cells[1].RowHeight = 56.26;
            Range collor = toyota.Range["A1:X1"];
            collor.Interior.Color = System.Drawing.Color.DarkBlue;
            collor.Font.Bold = true;
            collor.Font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            collor.HorizontalAlignment = Constants.xlCenter;
            collor.VerticalAlignment = Constants.xlCenter;

            #endregion

            toyota.Columns.EntireColumn.AutoFit();
            Range r2 = toyota.Rows[2];
            r2.EntireRow.Hidden = true;

            #region Borders
            Range table = toyota.Range["A1:X32"];

            table.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            table.Borders[XlBordersIndex.xlEdgeTop].Color = XlRgbColor.rgbBlack;
            table.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

            table.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            table.Borders[XlBordersIndex.xlEdgeLeft].Color = XlRgbColor.rgbBlack;
            table.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;

            table.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            table.Borders[XlBordersIndex.xlEdgeRight].Color = XlRgbColor.rgbBlack;
            table.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;

            table.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            table.Borders[XlBordersIndex.xlEdgeBottom].Color = XlRgbColor.rgbBlack;
            table.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

            table.HorizontalAlignment = Constants.xlCenter;
            table.VerticalAlignment = Constants.xlCenter;
            #endregion

            temp.Delete();
        }


        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Ocorreu um erro ao liberar o objeto do Excel: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
