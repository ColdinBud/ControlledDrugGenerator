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

                //Wait for next feature
                //string[] sourceDeviceList = Directory.GetFiles(Properties.Settings.Default["SourcePath"].ToString(),
                //    Properties.Settings.Default["DeviceFileName"].ToString() + "*" + fileDateFormat + "*.csv", SearchOption.TopDirectoryOnly);

                //string[] sourceLabelList = Directory.GetFiles(Properties.Settings.Default["SourcePath"].ToString(),
                //    Properties.Settings.Default["LabelFileName"].ToString() + "*" + fileDateFormat + "*.csv", SearchOption.TopDirectoryOnly);

                //將非注射用1-3級管制藥品取用紀錄移到紀錄表目錄底下
                MoveAndOrderFileList(sourceStationList);
                
                //將所有醫囑紀錄移到紀錄表目錄底下
                MoveAndOrderFileList(sourceOrderList);

                //Wait for next feature
                //將裝置記錄報表移到紀錄表目錄底下
                //MoveAndOrderFileList(sourceDeviceList);

                //將標籤記錄報表移到紀錄表目錄底下
                //MoveAndOrderFileList(sourceLabelList);
            }
            catch
            {
                result += "連接無效，無權限存取網路空間\r\n";
            }
            finally
            {
                //string stationFilePath = filePath + "\\" + Properties.Settings.Default["StationFileName"].ToString() + "*.csv";
                //string orderFilePath = filePath + "\\" + Properties.Settings.Default["OrderFileName"].ToString() + "*.csv";

                string nowTime = DateTime.Now.ToString("HH:mm");

                //Wait for next feature
                /*
                if (nowTime.Equals(Properties.Settings.Default.RunCompareTime))
                {
                    string[] deviceFilePathList = Directory.GetFiles(filePath,
                        Properties.Settings.Default.DeviceFileName + "*.csv", SearchOption.TopDirectoryOnly);
                    string deviceFilePath = deviceFilePathList.Length > 0 ? deviceFilePathList[0] : "";
                    string deviceFileName = deviceFilePath.Substring(deviceFilePath.LastIndexOf("\\") + 1);

                    string[] labelFilePathList = Directory.GetFiles(filePath,
                        Properties.Settings.Default.DeviceFileName + "*.csv", SearchOption.TopDirectoryOnly);
                    string labelFilePath = labelFilePathList.Length > 0 ? labelFilePathList[0] : "";
                    string labelFileName = labelFilePath.Substring(labelFilePath.LastIndexOf("\\") + 1);

                    if (!string.IsNullOrEmpty(deviceFileName) && !string.IsNullOrEmpty(labelFileName))
                    {

                    }
                }
                */

                //取出唯一的所有醫囑
                string[] orderFilePathList = Directory.GetFiles(filePath,
                    Properties.Settings.Default["OrderFileName"].ToString() + "*.csv", SearchOption.TopDirectoryOnly);
                string orderFilePath = orderFilePathList.Length > 0 ? orderFilePathList[0] : "";
                string orderFileName = orderFilePath.Substring(orderFilePath.LastIndexOf("\\") + 1);

                string[] stationFilePathList = Directory.GetFiles(filePath,
                    Properties.Settings.Default["StationFileName"].ToString() + "*.csv", SearchOption.TopDirectoryOnly);

                foreach (string stationFilePath in stationFilePathList)
                {
                    string stationFileName = stationFilePath.Substring(stationFilePath.LastIndexOf("\\") + 1);
                    if (!string.IsNullOrEmpty(orderFilePath) && !string.IsNullOrEmpty(stationFilePath))
                    {
                        List<ReportData> stationList = CombineCsv.CreateCsv(stationFilePath, orderFilePath);

                        MoveFile(stationFileName, stationFilePath, dateTime, nowTime);

                        result += new ExcelFormatter().CreateTotal(stationList, stationFileName);
                        if (stationList.Count > 0 && stationList[0].MedID != null)
                        {
                            result += new ExcelFormatter().FormatExcel(stationList, stationFileName);
                        }
                        else
                        {
                            result += "沒有更新資料\r\n";
                        }
                    }
                }

                if (!string.IsNullOrEmpty(orderFileName))
                {
                    MoveFile(orderFileName, orderFilePath, dateTime, nowTime);
                }

            }

            return result;
        }

        private static void MoveAndOrderFileList(string[] fileList)
        {
            string filePath = Properties.Settings.Default["FilePath"].ToString();
            string dateTime = DateTime.Now.ToString("yyyyMMdd");

            foreach (string str in fileList)
            {
                string fileName = str.Substring(str.LastIndexOf("\\") + 1);
                if (!File.Exists(string.Format("{0}\\{1}\\{3}\\{2}", filePath, dateTime, fileName, "原始資料")))
                {
                    if (!File.Exists(String.Format("{0}\\{2}", filePath, dateTime, fileName)))
                    {
                        File.Copy(str, String.Format("{0}\\{2}", filePath, dateTime, fileName), false);
                    }
                }
            }
        }

        private static string GetFileName(string fileTypeName)
        {
            string filePath = Properties.Settings.Default["FilePath"].ToString();

            string[] retFilePathList = Directory.GetFiles(filePath,
                    Properties.Settings.Default[fileTypeName].ToString() + "*.csv", SearchOption.TopDirectoryOnly);
            string retFilePath = retFilePathList.Length > 0 ? retFilePathList[0] : "";
            string retFileName = retFilePath.Substring(retFilePath.LastIndexOf("\\") + 1);

            return retFileName;
        }

        private static void MoveFile(string fileName, string retFilePath, string dateTime, string nowTime)
        {
            string filePath = Properties.Settings.Default["FilePath"].ToString();

            if (!File.Exists(filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + fileName))
            {
                System.IO.File.Move(retFilePath, filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + fileName);
            }
            else
            {
                System.IO.File.Move(retFilePath, filePath + "\\" + dateTime + "\\" + "原始資料" + "\\" + nowTime + fileName);
            }

        }
    }
}
