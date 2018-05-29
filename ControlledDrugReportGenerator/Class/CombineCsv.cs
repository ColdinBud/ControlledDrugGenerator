using ControlledDrugReportGenerator.Data;
using ServiceStack.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ControlledDrugReportGenerator.Class
{
    class CombineCsv
    {
        public static List<ReportData> CreateCsv(string stationFile, string orderFile)
        {
            var orderCsv = File.ReadAllText(orderFile);

            string[] propNames = null;
            List<string[]> rows = new List<string[]>();
            foreach (var line in CsvReader.ParseLines(orderCsv))
            {
                string[] strArray = CsvReader.ParseFields(line).ToArray();
                if (propNames == null)
                {
                    propNames = strArray;
                }
                else
                {
                    rows.Add(strArray);
                }
            }

            List<OrderData> orderList = new List<OrderData>();
            for (int r = 0; r < rows.Count; r++)
            {
                OrderData order = new OrderData();
                var cells = rows[r];
                for (int c = 0; c < cells.Length; c++)
                {
                    switch (c)
                    {
                        case 8:
                            order.Unit = cells[c];
                            break;
                        case 10:
                            order.OrderID = cells[c];
                            break;
                        case 20:
                            order.Dose = cells[c];
                            break;
                        default:
                            break;
                    }
                }
                orderList.Add(order);
            }

            var stationCsv = File.ReadAllText(stationFile);

            string[] stationPropNames = null;
            List<string[]> stationRows = new List<string[]>();
            foreach (var stationLine in CsvReader.ParseLines(stationCsv))
            {
                string[] strArray = CsvReader.ParseFields(stationLine).ToArray();
                if (stationPropNames == null)
                {
                    stationPropNames = strArray;
                }
                else
                {
                    stationRows.Add(strArray);
                }
            }
            List<ReportData> stationList = new List<ReportData>();
            for (int r = 0; r < stationRows.Count; r++)
            {
                ReportData report = new ReportData();
                var stationCells = stationRows[r];
                for (int c = 0; c < stationCells.Length; c++)
                {
                    switch (c)
                    {
                        case 0:
                            report.MedName = stationCells[c];
                            break;
                        case 1:
                            report.MedID = stationCells[c];
                            break;
                        case 5:
                            report.UsingUnit = stationCells[c];
                            break;
                        case 9:
                            report.Quantity = stationCells[c];
                            break;
                        case 10:
                            report.QuantityUnit = stationCells[c];
                            break;
                        case 12:
                            report.EndDose = stationCells[c];
                            break;
                        case 24:
                            report.UserName = stationCells[c];
                            break;
                        case 31:
                            report.OrderID = stationCells[c];
                            break;
                        case 34:
                            report.PatientID = stationCells[c];
                            break;
                        case 35:
                            report.PatientName = stationCells[c];
                            break;
                        case 37:
                            report.TransactionDate = stationCells[c];
                            break;
                        default:
                            break;
                    }
                }

                var findOrder = orderList.FirstOrDefault(o => o.OrderID == report.OrderID);
                if (findOrder != null)
                {
                    report.Dose = findOrder.Dose;
                }
                else
                {
                    report.Dose = "於醫囑內找不到對應資料";
                }
                stationList.Add(report);
            }

            return stationList;
        }
    }
}
