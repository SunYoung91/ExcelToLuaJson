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

                    var exporter_c = new Export( fileName, export_path.Text + "\\client\\", "c", "json", textLog);
                    exporter_c.DoExport();

                    var exporter_s = new Export( fileName, export_path.Text + "\\server\\", "s", "lua", textLog);
                    exporter_s.DoExport();

                }
            }
        }

        private void client_path_TextChanged(object sender, EventArgs e)
        {

        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }
    }
}
