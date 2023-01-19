using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelExport
{
    public partial class FormMain : Form
    {
        private string _currentDirectory = "";
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            _currentDirectory = Directory.GetCurrentDirectory();
            LogUtils.instance.SetTextBox(textLog);
            LoadConfig();
            RefreshXls();

        }

        private void RefreshXls()
        {
            String[] fileList = Directory.GetFiles(_currentDirectory, "*.xlsx");

            xlsFileList.Items.Clear();
            foreach (String filename in fileList)
            {
                xlsFileList.Items.Add(filename, false);
            }
        }

        private void LoadConfig()
        {
            var configFile = _currentDirectory + "\\config.txt";
            if (File.Exists(configFile))
            {
                StreamReader sr = new StreamReader(configFile, Encoding.Default);
                export_path.Text = sr.ReadLine();
                sr.Close();
            }
        }

        private void SaveConfig()
        {
            var configFile = _currentDirectory + "\\config.txt";
            var fileStream = new FileStream(configFile, FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(fileStream);
            sw.WriteLine(export_path.Text);
            sw.Flush();
            sw.Close();
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
     
            for (int i = 0; i < xlsFileList.Items.Count; i++)
            {
                if (xlsFileList.GetItemChecked(i))
                {
                    
                    var fileName = xlsFileList.GetItemText(xlsFileList.Items[i]);
                    LogUtils.instance.AddLog("读取Excel :" + fileName);
                    var exporter = new Export( fileName);
                    LogUtils.instance.AddLog("读取完成 : " + fileName);

                    LogUtils.instance.AddLog("开始导出到JSON");
                    exporter.DoExport(export_path.Text, "c", "json");

                    //LogUtils.instance.AddLog("开始导出到LUA");
                    //exporter.DoExport(export_path.Text + "\\data\\server\\", "s", "lua");

                    LogUtils.instance.AddLog("完成处理Excel :" + fileName);

                }
            }

            LogUtils.instance.AddLog("=======导出完成=====");
        }

        private void client_path_TextChanged(object sender, EventArgs e)
        {

        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }

        private void btn_select_all_Click(object sender, EventArgs e)
        {
            for  (var i = 0; i < xlsFileList.Items.Count; i++)
            {
                xlsFileList.SetItemChecked(i, true);
            }
        }

        private void btn_select_none_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < xlsFileList.Items.Count; i++)
            {
                xlsFileList.SetItemChecked(i, false);
            }
        }

        private void btn_exchange_select_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < xlsFileList.Items.Count; i++)
            {
                xlsFileList.SetItemChecked(i, !xlsFileList.GetItemChecked(i));
            }
        }
    }
}
