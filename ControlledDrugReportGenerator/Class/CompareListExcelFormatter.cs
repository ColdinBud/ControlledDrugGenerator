using ControlledDrugReportGenerator.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ControlledDrugReportGenerator.Class
{
    class CompareListExcelFormatter
    {
        public string CreateTotal(List<LabelData> labelList, string fileName = "")
        {
            string currentData = DateTime.Now.ToString("yyyyMMdd");
            string currentDateTime = DateTime.Now.ToString("HHmm");

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            Excel._Workbook wBook = excelApp.Workbooks.Add(Type.Missing);
            wBook.Activate();

            Excel._Worksheet wSheet = excelApp.ActiveSheet;
            wSheet.Name = "總表";
            Excel.Range wRange;

            var printPageSetup = wSheet.PageSetup;
            printPageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            printPageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            printPageSetup.FitToPagesWide = 1;

            excelApp.Cells[1, 2] = "分檢交貨與裝置庫存比較表";
            excelApp.Cells[2, 2] = "日期: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            if (labelList.Count > 0)
            {
                var nameGroup = from s in labelList
                                group s by new { s.Drawer, s.MedID, s.MedName, s.Min, s.Max, s.Current, s.Amount } into g
                                select new
                                {
                                    Drawer = g.Key.Drawer,
                                    MedID = g.Key.MedID,
                                    MedName = g.Key.MedName,
                                    Min = g.Key.Min,
                                    Max = g.Key.Max,
                                    Current = g.Key.Current,
                                    Amount = g.Key.Amount
                                };

                excelApp.Cells[2, 4] = "使用單位: " + fileName;
                //wSheet.get_Range("C2", "E2").Merge(wSheet.get_Range("C2", "E2").MergeCells);

                excelApp.Cells[3, 1] = "No.";
                excelApp.Cells[3, 2] = "藥品名稱";
                excelApp.Cells[3, 3] = "藥品八碼";
                excelApp.Cells[3, 4] = "藥格位置";
                excelApp.Cells[3, 5] = "最小值";
                excelApp.Cells[3, 6] = "最大值";
                excelApp.Cells[3, 7] = "庫存量";
                excelApp.Cells[3, 8] = "補藥量";

                int lineCount = 4;

                foreach (var ng in nameGroup)
                {
                    var groupData = labelList.Select(x => x.MedID == ng.MedID).ToArray();

                    excelApp.Cells[lineCount, 1] = (lineCount - 3);
                    excelApp.Cells[lineCount, 2] = ng.MedName;
                    excelApp.Cells[lineCount, 3] = ng.MedID;
                    excelApp.Cells[lineCount, 4] = ng.Drawer;
                    excelApp.Cells[lineCount, 5] = ng.Min;
                    excelApp.Cells[lineCount, 6] = ng.Max;
                    excelApp.Cells[lineCount, 7] = ng.Current;
                    excelApp.Cells[lineCount, 8] = ng.Amount;

                    lineCount++;
                }
            }
            else
            {
                excelApp.Cells[3, 1] = "無異常";
            }

            string areaName = "";
            if (!string.IsNullOrEmpty(fileName) && !string.IsNullOrEmpty(fileName.Split(' ')[0]))
            {
                areaName = fileName;
                //areaName = fileName.Split(' ')[0].Substring(16);
            }

            string mailName = currentData + "-" + currentDateTime + "-" + areaName + "-比對總表";
            string pathFile = $"{Properties.Settings.Default.FilePath}\\{currentData}\\{currentData}-{currentDateTime}-{areaName}-比對總表";

            wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            System.Threading.Thread.Sleep(5000);
            if (labelList.Count > 0)
            {
                SendEmail.SendMail(pathFile, mailName);
            }

            return "已建立 " + pathFile + "\r\n";
        }
    }
}
