using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelExport
{

    class LogUtils
    {
        public static LogUtils instance { get; private set; } = null;
        private TextBox _logBox;

        private LogUtils()
        {
           // Console.WriteLine("构造函数日志");
        }

        static LogUtils()
        {
            instance = new LogUtils();
        }

        public void SetTextBox(TextBox box)
        {
            _logBox = box;
        }

        public void AddLog(string log)
        {
            _logBox.AppendText(DateTime.Now.ToString() + "\t" + log + "\r\n");
        }

    }
}
