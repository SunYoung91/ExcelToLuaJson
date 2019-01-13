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
        private string _exportBasePath = "";
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            String rootPath = Directory.GetCurrentDirectory();
            String[] fileList =  Directory.GetFiles(rootPath, "*.xlsx");

            xlsFileList.Items.Clear();
            foreach(String filename in fileList)
            {
                xlsFileList.Items.Add(filename, false);
            }

        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            String rootPath = Directory.GetCurrentDirectory();
            for (int i = 0; i < xlsFileList.Items.Count; i++)
            {
                if (xlsFileList.GetItemChecked(i))
                {
                    var fileName = xlsFileList.GetItemText(xlsFileList.Items[i]);

                    var exporter_c = new Export(rootPath + fileName, _exportBasePath, "c", "json", textLog);
                    exporter_c.DoExport();

                    var exporter_s = new Export(rootPath + fileName, _exportBasePath, "s", "lua", textLog);
                    exporter_s.DoExport();

                }
            }
        }
    }
}
