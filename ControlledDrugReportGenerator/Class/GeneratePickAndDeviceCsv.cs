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
    class GeneratePickAndDeviceCsv
    {
        public static List<LabelData> ProcessLabelList(string[] csvList)
        {
            string result = "";

            int countCsv = 0;
            if (csvList.Length > 0)
            {
                string fileName = csvList[countCsv].Substring(csvList[0].LastIndexOf('\\') + 1);
                var csv = File.ReadAllText(csvList[countCsv]);
                result += $"資料處理中 - {csvList[countCsv]}\r\n";

                string[] propNames = null;
                List<string[]> rows = new List<string[]>();
                foreach (var line in CsvReader.ParseLines(csv))
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

                List<LabelData> labelList = new List<LabelData>();
                for (int r = 0; r < rows.Count; r++)
                {
                    LabelData label = new LabelData();
                    var cells = rows[r];
                    for (int c = 0; c < cells.Length; c++)
                    {
                        switch (c)
                        {
                            case 0:
                                label.Device = cells[c];
                                break;
                            case 1:
                                label.Drawer = cells[c];
                                break;
                            case 2:
                                label.MedID = cells[c];
                                break;
                            case 3:
                                label.MedName = cells[c];
                                break;
                            case 5:
                                label.Min = Int32.Parse(cells[c]);
                                break;
                            case 6:
                                label.Max = Int32.Parse(cells[c]);
                                break;
                            case 7:
                                label.Current = Int32.Parse(cells[c]);
                                break;
                            case 12:
                                label.Amount = Int32.Parse(cells[c]);
                                break;
                            default:
                                break;
                        }
                    }

                    if (!(string.IsNullOrEmpty(label.MedID)))
                    {
                        labelList.Add(label);
                    }
                }

                return labelList;
            }
            return null;
        }

        public static List<LabelData> ProcessDeviceList(string[] csvList)
        {
            int countCsv = 0;
            if (csvList.Length > 0)
            {
                string fileName = csvList[countCsv].Substring(csvList[countCsv].LastIndexOf('\\') + 1);
                var csv = File.ReadAllText(csvList[countCsv]);

                string[] propNames = null;
                List<string[]> rows = new List<string[]>();
                foreach (var line in CsvReader.ParseLines(csv))
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

                List<LabelData> labelList = new List<LabelData>();
                for (int r = 0; r < rows.Count; r++)
                {
                    LabelData label = new LabelData();
                    var cells = rows[r];
                    for (int c = 0; c < cells.Length; c++)
                    {
                        switch (c)
                        {
                            case 0:
                                label.MedName = cells[c];
                                break;
                            case 2:
                                label.MedID = cells[c];
                                break;
                            case 7:
                                label.Device = cells[c];
                                break;
                            case 8:
                                label.Drawer = cells[c];
                                break;
                            case 9:
                                label.Current = Int32.Parse(cells[c]);
                                break;
                            case 10:
                                label.Min = Int32.Parse(cells[c]); ;
                                break;
                            case 11:
                                label.Max = Int32.Parse(cells[c]); ;
                                break;
                            default:
                                break;
                        }
                    }

                    if (IsMaxDay())
                    {
                        label.Amount = (label.Max > label.Current) ? (label.Max - label.Current) : 0;
                    }
                    else
                    {
                        label.Amount = (label.Min >= label.Current) ? (label.Max - label.Current) : 0;
                    }

                    if (!(string.IsNullOrEmpty(label.MedID)))
                    {
                        labelList.Add(label);
                    }

                }

                if (!IsMaxDay())
                {
                    foreach (var v in labelList)
                    {
                        int totalMin = labelList.Where(x => x.MedID == v.MedID).Sum(x => x.Min);
                        int totalMax = labelList.Where(x => x.MedID == v.MedID).Sum(x => x.Max);
                        int totalCurrent = labelList.Where(x => x.MedID == v.MedID).Sum(x => x.Current);

                        if (totalMin < totalCurrent)
                        {
                            labelList = labelList.Where(x => x.MedID != v.MedID).ToList();
                        }

                    }
                }

                labelList = labelList.Where(x => x.Amount != 0).ToList();

                return labelList;
            }

            return null;
        }

        private static bool IsMaxDay()
        {
            string[] maxDay = Properties.Settings.Default.MaxDay.Split(',');

            foreach (var s in maxDay)
            {
                DayOfWeek d = (DayOfWeek)Int32.Parse(s);
                if (DateTime.Today.DayOfWeek == d)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
