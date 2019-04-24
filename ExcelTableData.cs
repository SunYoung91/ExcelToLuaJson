using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace ExcelExport
{
    public class ExcelTableData
    {
        public List<ExcelSheetData> excelSheets = new List<ExcelSheetData>();

        public void LoadFromFile(string fileName)
        {
            var stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);         
            for (int i = 0; i < reader.ResultsCount; i++)
            {            
                var sheet = new ExcelSheetData();
                LogUtils.instance.AddLog("读取页签:" + reader.Name);
                sheet.ReadFromExcel(reader.Name, reader);

                if (sheet.isNeedExprot)
                {               
                    excelSheets.Add(sheet);
                } else
                {
                    LogUtils.instance.AddLog("页签无需导出 , 跳过 :" + reader.Name);
                }
                  
                reader.NextResult();
            }

            stream.Close();
        }
    }
}
