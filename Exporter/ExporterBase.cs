using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelExport.Exporter
{
    public class ExporterBase
    {

        protected virtual void ExportTiny(ExcelSheetData data,StreamWriter writer)
        {
            //nothing
        }

        protected virtual void ExportBase(ExcelSheetData data,StreamWriter writer)
        {
            //nothing
        }


        protected virtual void AddHeader(ExcelSheetData data, StreamWriter writer)
        {
            writer.WriteLine(data.exportHeader);
        }

        protected virtual void AddEnd(ExcelSheetData data, StreamWriter writer)
        {
            writer.WriteLine(data.exportEnd);
        }

        //子类只需按情况实现上面4个函数。


        private void CheckCreateDir(string dir)
        {
            var targetDir = Directory.GetParent(dir).ToString();
            if (!Directory.Exists(targetDir))
            {        
                Directory.CreateDirectory(targetDir);
            }
        }

        public void SaveToFile(ExcelSheetData data , string fileName)
        {

            CheckCreateDir(fileName);

            var stream = new FileStream(fileName, FileMode.Create);
            var writer = new StreamWriter(stream);

            AddHeader(data, writer);
            if (data.exportSchema == "base")
            {
                ExportBase(data, writer);
            } else if(data.exportSchema == "tiny")
            {
                ExportTiny(data, writer);
            }

            AddEnd(data, writer);
            writer.Close();
            stream.Close();

        }

    }
}
