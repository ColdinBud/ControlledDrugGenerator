using ControlledDrugReportGenerator.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ControlledDrugReportGenerator.Class
{
    class DoService
    {
        public static string pubService = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString();
        public static string pubVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        public static string printResult()
        {
            string result = "";

            string filePath = Properties.Settings.Default["FilePath"].ToString();

            string dateTime = DateTime.Now.ToString("yyyyMMdd");
            string fileDateFormat = DateTime.Now.ToString("M-d-yyyy");

            Directory.CreateDirectory(String.Format("{0}\\{1}", filePath, dateTime));
            Directory.CreateDirectory(String.Format("{0}\\{1}\\{2}", filePath, dateTime, "原始資料"));

            try
            {
                string[] sourceStationList = Directory.GetFiles(Properties.Settings.Default["SourcePath"].ToString(),
                    Properties.Settings.Default["StationFileName"].ToString() + "*" + fileDateFormat + "*.csv", SearchOption.TopDirectoryOnly);

                string[] sourceOrderList = Directory.GetFiles(Properties.Settings.Default["SourcePath"].ToString(),
                    Properties.Settings.Default["OrderFileName"].ToString() + "*" + fileDateFormat + "*.csv", SearchOption.TopDirectoryOnly);

                foreach (string str in sourceStationList)
                {
                    string stationFileName = str.Substring(str.LastIndexOf("\\") + 1);
                    if (!File.Exists(string.Format("{0}\\{1}\\{3}\\{2}", filePath, dateTime, stationFileName, "原始資料")))
                    {
                        if (!File.Exists(String.Format("{0}\\{2}", filePath, dateTime, stationFileName)))
                        {
                            File.Copy(str, String.Format("{0}\\{2}", filePath, dateTime, stationFileName), false);
                        }
                    }
                }

                foreach (string str in sourceOrderList)
                {
                    string orderFileName = str.Substring(str.LastIndexOf("\\") + 1);
                    if (!File.Exists(string.Format("{0}\\{1}\\{3}\\{2}", filePath, dateTime, orderFileName, "原始資料")))
                    {
                        if (!File.Exists(String.Format("{0}\\{2}", filePath, dateTime, orderFileName)))
                        {
                            File.Copy(str, String.Format("{0}\\{2}", filePath, dateTime, orderFileName), false);
                        }
                    }
                }
            }
            catch
            {
                result += "連接無效，無權限存取網路空間\r\n";
            }
            finally
            {
                //string stationFilePath = filePath + "\\" + Properties.Settings.Default["StationFileName"].ToString() + "*.csv";
                //string orderFilePath = filePath + "\\" + Properties.Settings.Default["OrderFileName"].ToString() + "*.csv";

                string[] stationFilePathList = Directory.GetFiles(filePath,
                    Properties.Settings.Default["StationFileName"].ToString() + "*.csv", SearchOption.TopDirectoryOnly);
                string stationFilePath = stationFilePathList.Length > 0 ? stationFilePathList[0] : "";

                string[] orderFilePathList = Directory.GetFiles(filePath,
                    Properties.Settings.Default["OrderFileName"].ToString() + "*.csv", SearchOption.TopDirectoryOnly);
                string orderFilePath = orderFilePathList.Length > 0 ? orderFilePathList[0] : "";

                string stationFileName = stationFilePath.Substring(stationFilePath.LastIndexOf("\\") + 1);
                string orderFileName = orderFilePath.Substring(orderFilePath.LastIndexOf("\\") + 1);

                if (!string.IsNullOrEmpty(orderFilePath) && !string.IsNullOrEmpty(stationFilePath))
                {
                    List<ReportData> stationList = CombineCsv.CreateCsv(stationFilePath, orderFilePath);

                    //move source file
                    if (!File.Exists(filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + stationFileName))
                    {
                        System.IO.File.Move(stationFilePath, filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + stationFileName);
                    }
                    else
                    {
                        string nowTime = DateTime.Now.ToString("HHmm");
                        System.IO.File.Move(stationFilePath, filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + nowTime + stationFileName);
                    }
                    if (!File.Exists(filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + orderFileName))
                    {
                        System.IO.File.Move(orderFilePath, filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + orderFileName);
                    }
                    else
                    {
                        string nowTime = DateTime.Now.ToString("HHmm");
                        System.IO.File.Move(orderFilePath, filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + nowTime + orderFileName);
                    }

                    result += new ExcelFormatter().CreateTotal(stationList);
                    result += new ExcelFormatter().FormatExcel(stationList);
                }

            }

            return result;
        }
    }
}
