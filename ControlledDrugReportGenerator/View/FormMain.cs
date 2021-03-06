﻿using ControlledDrugReportGenerator.Class;
using ControlledDrugReportGenerator.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ControlledDrugReportGenerator.View
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        Thread combineListThread;

        private void cfgButton_Click(object sender, EventArgs e)
        {
            FormConfig objConfig = new FormConfig();
            if (objConfig.ShowDialog(this) == DialogResult.OK)
            {

            }
            objConfig.Dispose();
        }

        private void startBtn_Click(object sender, EventArgs e)
        {
            timerLog.Interval = Properties.Settings.Default.Interval * 1000;
            timerLog.Enabled = true;
            startBtn.Enabled = false;
            stopBtn.Enabled = true;

            txtMessage.Text += "排程已啟動.....\r\n";
        }

        private void stopBtn_Click(object sender, EventArgs e)
        {
            timerLog.Enabled = false;
            startBtn.Enabled = true;
            stopBtn.Enabled = false;
        }

        private void timerLog_Tick(object sender, EventArgs e)
        {
            if (txtMessage.Lines.Length > 14)
            {
                txtMessage.Clear();
            }
            txtMessage.Text += string.IsNullOrEmpty(txtMessage.Text) ? "\r\n" : "";
            //txtMessage.Text += $"{DateTime.Now.ToString("HH:mm:ss")}:    資料處理中，請稍後...\r\n";

            string printResult = DoService.printResult();
            string res = string.IsNullOrEmpty(printResult) ? "沒有新資料" : printResult;
            string test = $"{DateTime.Now.ToString("HH:mm:ss")}:    {res}\r\n";

            string msg = test;

            txtMessage.Text += msg;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "請選擇取藥紀錄";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "csv files (*.*)|*.csv";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                label1.Text = dialog.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "請選擇所有醫囑";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "csv files (*.*)|*.csv";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                label2.Text = dialog.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            stopBtn_Click(sender, e);

            string dateFormat = DateTime.Now.ToString("yyyyMMdd");
            string dateTime = DateTime.Now.ToString("HHmm");

            string stationFileName = label1.Text.Substring(label1.Text.LastIndexOf("\\") + 1);
            string orderFileName = label2.Text.Substring(label2.Text.LastIndexOf("\\") + 1);

            Directory.CreateDirectory($"{Properties.Settings.Default.FilePath}\\{dateFormat}\\原始資料");

            File.Copy(label1.Text, $"{Properties.Settings.Default.FilePath}\\{dateFormat}\\原始資料\\{dateTime}-{stationFileName}", false);
            File.Copy(label2.Text, $"{Properties.Settings.Default.FilePath}\\{dateFormat}\\原始資料\\{dateTime}-{orderFileName}", false);

            List<ReportData> stationList = CombineCsv.CreateCsv(label1.Text, label2.Text);

            string currentTime = DateTime.Now.ToString("HH:mm:ss");

            txtMessage.Text += $"{currentTime}:    資料處理中，請稍後...\r\n";
            txtMessage.Text += $"{currentTime}:    {new ExcelFormatter().CreateTotal(stationList)}";
            if (stationList.Count > 0 && stationList[0].MedID != null)
            {
                txtMessage.Text += $"{currentTime}:    {new ExcelFormatter().FormatExcel(stationList)}";
            }
            else
            {
                txtMessage.Text += "沒有更新資料\r\n";
            }
            //button3.Enabled = true;
        }
    }
}
