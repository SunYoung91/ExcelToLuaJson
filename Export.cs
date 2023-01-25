using System;
using ExcelDataReader;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using ExcelExport.Exporter;

namespace ExcelExport
{
    class Export
    {
        private string _ExportExcelFileName = "";
        private ExcelTableData excelTable;
        public Export(string excelFileName)
        {
            _ExportExcelFileName = excelFileName;
            excelTable = new ExcelTableData();
            excelTable.LoadFromFile(excelFileName);
        }

        public void DoExport(string exportBasePath, string exportMode, string exportType)
        {
            ExporterBase exporter = null;
            switch (exportType)
            {
                case "lua":
                    {
                        exporter = new ExporterLua(exportMode);
                    }
                    break;
                case "json":
                    {
                        exporter = new ExporterJson(exportMode);
                    }
                    break;
            }


            for (var i = 0; i < excelTable.excelSheets.Count; i++)
            {
                LogUtils.instance.AddLog("导出页签 : " + excelTable.excelSheets[i].sheetName);
                exporter.SaveToFile(excelTable.excelSheets[i], exportBasePath + "\\" + excelTable.excelSheets[i].exportPath);
            }

        }


    }
}
