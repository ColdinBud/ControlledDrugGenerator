using ControlledDrugReportGenerator.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ControlledDrugReportGenerator.Class
{
    class ExcelFormatter
    {
        public string CreateTotal(List<ReportData> stationList)
        {
            var nameGroup = from s in stationList
                            group s by new { s.MedID, s.MedName, s.QuantityUnit } into g
                            select new
                            {
                                MedID = g.Key.MedID,
                                MedName = g.Key.MedName,
                                Unit = g.Key.QuantityUnit,
                                Num = g.Count(),
                                Total = g.Sum(s => int.Parse(s.Quantity))
                            };

            string currentDate = DateTime.Now.ToString("yyyyMMdd");
            string currentDateTime = DateTime.Now.ToString("HHmm");

            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;


            excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            wBook = excelApp.Workbooks.Add(Type.Missing);
            wBook.Activate();
            wSheet = excelApp.ActiveSheet;
            wSheet.Name = "總表";

            var printPageSetup = wSheet.PageSetup;
            printPageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            printPageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            excelApp.Cells[1, 1] = "No.";
            excelApp.Cells[1, 2] = "藥品名稱";
            excelApp.Cells[1, 3] = "藥品八碼";
            excelApp.Cells[1, 4] = "取藥筆數";
            excelApp.Cells[1, 5] = "總取用量";
            excelApp.Cells[1, 6] = "取用單位";

            int lineCount = 2;
            foreach (var ng in nameGroup)
            {
                var groupData = (from g in stationList
                                 where g.MedID == ng.MedID
                                 select g).ToArray();

                excelApp.Cells[lineCount, 1] = (lineCount - 1);
                excelApp.Cells[lineCount, 2] = ng.MedName;
                excelApp.Cells[lineCount, 3] = ng.MedID;
                excelApp.Cells[lineCount, 4] = ng.Num;
                excelApp.Cells[lineCount, 5] = ng.Total;
                excelApp.Cells[lineCount, 6] = ng.Unit;

                lineCount++;

                foreach (var group in groupData)
                {
                    if (group.Dose.Equals("於醫囑內找不到對應資料"))
                    {
                        excelApp.Cells[lineCount, 1] = (lineCount - 1);
                        excelApp.Cells[lineCount, 2] = ng.MedID + "-" +group.OrderID +　" - 找不到對應的醫囑資料";
                        lineCount++;
                    }
                }
            }

            wSheet.get_Range($"A1", $"F{lineCount - 1}").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[lineCount - 1, 6]];

            wRange.Columns.AutoFit();

            string pathFile = $"{Properties.Settings.Default.FilePath}\\{currentDate}\\{currentDate}-{currentDateTime}-總表";

            wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wSheet.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Properties.Settings.Default.ActivePrinter,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            /*
            bool userDidntCancel = excelApp.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogPrint].Show(Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            */
            

            wBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            return "已建立 " + pathFile + "\r\n";
        }

        public string FormatExcel(List<ReportData> stationList)
        {
            string result = "";

            var nameGroup = from s in stationList
                            group s by new { s.MedID, s.QuantityUnit } into g
                            select new { MedID = g.Key.MedID, Total = g.Sum(s => int.Parse(s.Quantity)) };


            string currentDate = DateTime.Now.ToString("yyyyMMdd");
            string currentDateTime = DateTime.Now.ToString("HHmm");

            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            wBook = excelApp.Workbooks.Add(Type.Missing);
            //excelApp.Workbooks.Add(Type.Missing);
            //wBook = excelApp.Workbooks[nameGroup.Count()];
            wBook.Activate();

            foreach (var ng in nameGroup)
            {
                //Console.WriteLine(group.MedID);
                var groupData = (from g in stationList
                                 where g.MedID == ng.MedID
                                 select g).ToArray();


                wSheet = (Excel._Worksheet)wBook.Worksheets.Add();
                wSheet.Name = groupData[0].MedID;
                //wSheet.Activate();

                var printPageSetup = wSheet.PageSetup;
                printPageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                printPageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                printPageSetup.FitToPagesWide = 1;
                printPageSetup.FitToPagesTall = 1;
                printPageSetup.LeftMargin = excelApp.InchesToPoints(0.7);
                printPageSetup.RightMargin = excelApp.InchesToPoints(0.3);
                printPageSetup.TopMargin = excelApp.InchesToPoints(0.75);
                printPageSetup.BottomMargin = excelApp.InchesToPoints(0.7);
                printPageSetup.HeaderMargin = excelApp.InchesToPoints(0.3);
                printPageSetup.FooterMargin = excelApp.InchesToPoints(0.3);

                int curPage = 1;
                int curLine = 1;
                while (curPage * 10 <= groupData.Count())
                {
                    curLine = (curPage - 1) * 47 + 1;
                    wSheet.get_Range($"A{curLine + 4}", $"D{curLine + 4}").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelApp.Cells[curLine + 4, 1] = "No.";
                    excelApp.Cells[curLine + 4, 2] = "病人基本資料";
                    excelApp.Cells[curLine + 4, 3] = "處方資料及使用紀錄";
                    excelApp.Cells[curLine + 4, 4] = "取用量";
                    //excelApp.Cells[curLine + 4, 5] = "結存量";

                    string RangeCenter = $"A{curLine}:D{curLine + 1}";
                    wSheet.get_Range(RangeCenter).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelApp.Cells[curLine, 1] = "國立臺灣大學醫學院附設醫院";
                    wSheet.get_Range($"A{curLine}", $"D{curLine}").Merge(wSheet.get_Range($"A{curLine}", $"D{curLine}").MergeCells);

                    excelApp.Cells[curLine + 1, 1] = "非注射用1-3級管制藥品使用紀錄表";
                    wSheet.get_Range($"A{curLine + 1}", $"D{curLine + 1}").Merge(wSheet.get_Range($"A{curLine + 1}", $"D{curLine + 1}").MergeCells);

                    excelApp.Cells[curLine + 2, 1] = "使用單位：" + groupData[0].UsingUnit;
                    wSheet.get_Range($"A{curLine + 2}", $"D{curLine + 2}").Merge(wSheet.get_Range($"A{curLine + 2}", $"D{curLine + 2}").MergeCells);

                    excelApp.Cells[curLine + 3, 1] = "藥品名稱：" + groupData[0].MedName;
                    wSheet.get_Range($"A{curLine + 3}", $"D{curLine + 3}").Merge(wSheet.get_Range($"A{curLine + 3}", $"D{curLine + 3}").MergeCells);

                    int count = 1;
                    for (int i = 5; i < 45; i += 4)
                    {
                        string[] dateTime = groupData[((curPage - 1) * 10 + count) - 1].OrderStartTime.Split(' ')[0].Split('/');
                        string dateFormat = dateTime[0] + (dateTime[1].Length == 1 ? "0" + dateTime[1] : dateTime[1]) +
                            (dateTime[2].Length == 1 ? "0" + dateTime[2] : dateTime[2]);

                        string dosage = Regex.Match(groupData[((curPage - 1) * 10 + count) - 1].Dose, @"-?\d+(?:\.\d+)?").ToString();
                        string orderUnit = groupData[((curPage - 1) * 10) + count - 1].Dose.Substring(dosage.Length);

                        string orderID = groupData[((curPage - 1) * 10 + count) - 1].OrderID.Split('-')[0];

                        excelApp.Cells[curLine + i, 1] = (curPage - 1) * 10 + count;
                        excelApp.Cells[curLine + i, 2] = "病歷號碼：" + groupData[((curPage - 1) * 10 + count) - 1].PatientID;
                        excelApp.Cells[curLine + i, 3] = "處方號碼：" + dateFormat + "-" + orderID + "-" + groupData[((curPage - 1) * 10 + count) - 1].PatientID;
                        excelApp.Cells[curLine + i, 4] = groupData[((curPage - 1) * 10 + count) - 1].Quantity;
                        //excelApp.Cells[curLine + i, 5] = groupData[((curPage - 1) * 10 + count) - 1].EndDose;

                        excelApp.Cells[curLine + i + 1, 1] = "";
                        excelApp.Cells[curLine + i + 1, 2] = "病人姓名：" + groupData[((curPage - 1) * 10 + count) - 1].PatientName;
                        excelApp.Cells[curLine + i + 1, 3] = "使用日期：" + groupData[((curPage - 1) * 10 + count) - 1].TransactionDate;
                        excelApp.Cells[curLine + i + 1, 4] = "";
                        //excelApp.Cells[curLine + i + 1, 5] = "";

                        excelApp.Cells[curLine + i + 2, 1] = "";
                        excelApp.Cells[curLine + i + 2, 2] = "";
                        excelApp.Cells[curLine + i + 2, 3] = "處方劑量：" + dosage + " " + orderUnit;
                        excelApp.Cells[curLine + i + 2, 4] = "";
                        //excelApp.Cells[curLine + i + 2, 5] = "";

                        excelApp.Cells[curLine + i + 3, 1] = "";
                        excelApp.Cells[curLine + i + 3, 2] = "";
                        excelApp.Cells[curLine + i + 3, 3] = "領藥者：" + groupData[((curPage - 1) * 10 + count) - 1].UserName.Replace(", ", "");
                        excelApp.Cells[curLine + i + 3, 4] = "";
                        //excelApp.Cells[curLine + i + 3, 5] = "";

                        /*
                        wSheet.get_Range("A" + (curLine + i).ToString(), "A" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                            Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        wSheet.get_Range("B" + (curLine + i).ToString(), "B" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                            Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        wSheet.get_Range("C" + (curLine + i).ToString(), "C" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                            Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        wSheet.get_Range("D" + (curLine + i).ToString(), "D" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                            Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        wSheet.get_Range("E" + (curLine + i).ToString(), "E" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                           Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                           */

                        for (int col = 0; col < 4; col++)
                        {
                            string colcount = ((char)('A' + col)).ToString();

                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            wSheet.get_Range(colcount + (curLine + i + 3).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            wSheet.get_Range(colcount + (curLine + i + 3).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        }

                        count++;
                    }

                    excelApp.Cells[curLine + 46, 1] = "";
                    excelApp.Cells[curLine + 46, 2] = "";
                    excelApp.Cells[curLine + 46, 3] = "";
                    //int pageSetup = wSheet.PageSetup.Pages.Count;
                    excelApp.Cells[curLine + 46, 4] = $"第{curPage}頁, 共{groupData.Count() / 10 + 1}頁";
                    wSheet.get_Range($"C{curLine + 46}", $"D{curLine + 46}").Merge(wSheet.get_Range($"C{curLine + 46}", $"D{curLine + 46}").MergeCells);
                    wSheet.get_Range($"C{curLine + 46}:D{curLine + 46}").HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                    curPage++;
                }

                if (groupData.Count() % 10 != 0)
                {
                    curLine = (curPage - 1) * 47 + 1;
                    //wSheet.get_Range("A5", "E5").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    wSheet.get_Range($"A{curLine + 4}", $"D{curLine + 4}").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelApp.Cells[curLine + 4, 1] = "No.";
                    excelApp.Cells[curLine + 4, 2] = "病人基本資料";
                    excelApp.Cells[curLine + 4, 3] = "處方資料及使用紀錄";
                    excelApp.Cells[curLine + 4, 4] = "取用量";
                    //excelApp.Cells[curLine + 4, 5] = "結存量";

                    string rcenter = $"A{curLine}:D{curLine + 1}";
                    wSheet.get_Range(rcenter).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelApp.Cells[curLine, 1] = "國立臺灣大學醫學院附設醫院";
                    wSheet.get_Range($"A{curLine}", $"D{curLine}").Merge(wSheet.get_Range($"A{curLine}", $"D{curLine}").MergeCells);

                    excelApp.Cells[curLine + 1, 1] = "非注射用1-3級管制藥品使用紀錄表";
                    wSheet.get_Range($"A{curLine + 1}", $"D{curLine + 1}").Merge(wSheet.get_Range($"A{curLine + 1}", $"D{curLine + 1}").MergeCells);

                    excelApp.Cells[curLine + 2, 1] = "使用單位：" + groupData[0].UsingUnit;
                    wSheet.get_Range($"A{curLine + 2}", $"D{curLine + 2}").Merge(wSheet.get_Range($"A{curLine + 2}", $"D{curLine + 2}").MergeCells);

                    excelApp.Cells[curLine + 3, 1] = "藥品名稱：" + groupData[0].MedName;
                    wSheet.get_Range($"A{curLine + 3}", $"D{curLine + 3}").Merge(wSheet.get_Range($"A{curLine + 3}", $"D{curLine + 3}").MergeCells);

                    int listCount = 1;
                    int itemCount = groupData.Count() % 10;
                    for (int i = 5; i < itemCount * 4 + 5; i += 4)
                    {
                        string[] dateTime = groupData[((curPage - 1) * 10 + listCount) - 1].OrderStartTime.Split(' ')[0].Split('/');
                        string dateFormat = dateTime[0] + (dateTime[1].Length == 1 ? "0" + dateTime[1] : dateTime[1]) +
                            (dateTime[2].Length == 1 ? "0" + dateTime[2] : dateTime[2]);

                        string dosage = Regex.Match(groupData[((curPage - 1) * 10 + listCount) - 1].Dose, @"-?\d+(?:\.\d+)?").ToString();
                        string orderUnit = groupData[((curPage - 1) * 10) + listCount - 1].Dose.Substring(dosage.Length);

                        string orderID = groupData[((curPage - 1) * 10 + listCount) - 1].OrderID.Split('-')[0];

                        excelApp.Cells[curLine + i, 1] = (curPage - 1) * 10 + listCount;
                        excelApp.Cells[curLine + i, 2] = "病歷號碼：" + groupData[((curPage - 1) * 10 + listCount) - 1].PatientID;
                        excelApp.Cells[curLine + i, 3] = "處方號碼：" + dateFormat + "-" + orderID + "-" + groupData[((curPage - 1) * 10 + listCount) - 1].PatientID;
                        excelApp.Cells[curLine + i, 4] = groupData[((curPage - 1) * 10 + listCount) - 1].Quantity;
                        //excelApp.Cells[curLine + i, 5] = groupData[((curPage - 1) * 10 + listCount) - 1].EndDose;

                        excelApp.Cells[curLine + i + 1, 1] = "";
                        excelApp.Cells[curLine + i + 1, 2] = "病人姓名：" + groupData[((curPage - 1) * 10 + listCount) - 1].PatientName;
                        excelApp.Cells[curLine + i + 1, 3] = "使用日期：" + groupData[((curPage - 1) * 10 + listCount) - 1].TransactionDate;
                        excelApp.Cells[curLine + i + 1, 4] = "";
                        //excelApp.Cells[curLine + i + 1, 5] = "";

                        excelApp.Cells[curLine + i + 2, 1] = "";
                        excelApp.Cells[curLine + i + 2, 2] = "";
                        excelApp.Cells[curLine + i + 2, 3] = "處方劑量：" + dosage + " " + orderUnit;
                        excelApp.Cells[curLine + i + 2, 4] = "";
                        //excelApp.Cells[curLine + i + 2, 5] = "";

                        excelApp.Cells[curLine + i + 3, 1] = "";
                        excelApp.Cells[curLine + i + 3, 2] = "";
                        excelApp.Cells[curLine + i + 3, 3] = "領藥者：" + groupData[((curPage - 1) * 10 + listCount) - 1].UserName.Replace(", ", "");
                        excelApp.Cells[curLine + i + 3, 4] = "";
                        //excelApp.Cells[curLine + i + 3, 5] = "";


                        //wSheet.get_Range("A" + (curLine + i).ToString(), "A" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                        //    Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        //wSheet.get_Range("B" + (curLine + i).ToString(), "B" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                        //    Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        //wSheet.get_Range("C" + (curLine + i).ToString(), "C" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                        //    Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        //wSheet.get_Range("D" + (curLine + i).ToString(), "D" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                        //   Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                        //wSheet.get_Range("E" + (curLine + i).ToString(), "E" + (curLine + i + 3)).BorderAround2(Excel.XlLineStyle.xlContinuous,
                        //    Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);

                        //cellRange = oSheet.get_Range(oSheet.Cells[3, 10], oSheet.Cells[15, 11]);
                        //cellRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        //cellRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                        
                        for (int col = 0; col < 4; col++)
                        {
                            string colcount = ((char)('A' + col)).ToString();

                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            wSheet.get_Range(colcount + (curLine + i).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            wSheet.get_Range(colcount + (curLine + i + 3).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            wSheet.get_Range(colcount + (curLine + i + 3).ToString(), colcount + (curLine + i + 3)).
                                Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                        }

                        listCount++;
                    }
                }

                //decimal sum = groupData.AsEnumerable().Sum(s => Decimal.Parse(s.Quantity));

                curLine = groupData.Count() % 10 == 0 ? curLine + 47 : curLine + ((groupData.Count() % 10) * 4 + 5);
                //curLine = curLine + ((groupData.Count() % 10) * 4 + 5);
                excelApp.Cells[curLine + 1, 1] = "補藥量：" + ng.Total + " " + groupData[0].QuantityUnit;
                wSheet.get_Range($"A{curLine + 2}", $"B{curLine + 2}").Merge(wSheet.get_Range($"A{curLine + 2}", $"B{curLine + 2}").MergeCells);

                //excelApp.Cells[curLine + 1, 3] = "護理長/用藥監督人：";
                excelApp.Cells[curLine + 3, 1] = "藥師：";
                excelApp.Cells[curLine + 3, 3] = "調劑日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm"); ;
                excelApp.Cells[curLine + 3, 4] = "領藥人：";
                excelApp.Cells[curLine + 5, 4] = $"第{curPage}頁, 共{groupData.Count() / 10 + 1}頁";
                wSheet.get_Range($"C{curLine + 5}", $"D{curLine + 5}").Merge(wSheet.get_Range($"C{curLine + 5}", $"D{curLine + 5}").MergeCells);
                wSheet.get_Range($"C{curLine + 5}:D{curLine + 5}").HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[curLine + 6, 5]];
                wRange.Columns.AutoFit();

                //string savePdfPath = $"{Properties.Settings.Default.FilePath}\\{currentDate}\\{currentDate}-{currentDateTime}{groupData[0].MedID}";


                wSheet.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Properties.Settings.Default.ActivePrinter,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                //wSheet.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //wSheet.PageSetup.Application.ActivePrinter = Properties.Settings.Default.ActivePrinter;

                //wSheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, savePdfPath);
                //result += "已建立 " + savePdfPath + "\r\n";

                wSheet = null;
            }

            string pathFile = $"{Properties.Settings.Default.FilePath}\\{currentDate}\\{currentDate}-{currentDateTime}-非注射用1-3級管制藥品使用紀錄";

            wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            result += "已建立 " + pathFile + "\r\n";

            wBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            return result;
        }
    }
}
