using System;
using ExcelDataReader;
using System.IO;
using System.Windows.Forms;

namespace ExcelExport
{
    class Export
    {
        private string _ExportExcelFileName = "";
        private string _ExportBastPath = "";
        private string _ExportMode = "";//导出模式
        private string _ExportType = "json"; //导出模式
        private TextBox _LogBox = null;
        public Export(string excelFileName, string exportBasePath, string exportMode, string exportType, TextBox logBox)
        {
            _ExportExcelFileName = excelFileName;
            _ExportBastPath = exportBasePath;
            _ExportMode = exportMode;
            _ExportType = exportType;
            _LogBox = logBox;

        }

        public void DoExport()
        {
            var stream = File.Open(_ExportExcelFileName, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            AddLog(string.Format("文件: {0} , 页签数量: {1} ", _ExportExcelFileName, reader.ResultsCount));
            for (int i = 0; i < reader.ResultsCount; i++)
            {
                ExportSheet(reader);
                reader.NextResult();
            }

            stream.Close();
        }

        private void AddLog(string log)
        {
            _LogBox.AppendText(DateTime.Now.ToString() + "\t" + log + "\r\n");
        }

        private void ExportSheet(IExcelDataReader dataTable)
        {
            AddLog(string.Format("页签: {0} , 行数量: {1} , 列数量:{2} " , dataTable.Name, dataTable.RowCount, dataTable.FieldCount));
            var exportSchema = "none"; //导出类型
            var exportPath = ""; //导出路径
            var keyCount = 0;  //数组层级数量
            var exportHeader = ""; //导出文件头
            var exportEnd = "";//导出文件尾

            //建立所有会缓存的数组
            var fieldDatas = new FieldData[dataTable.FieldCount];
            for (int i = 0; i < dataTable.FieldCount; i++)
            {
                fieldDatas[i] = new FieldData();
            }

            //迭代读取所有数据
            for (int rowIndex = 0; rowIndex < dataTable.RowCount; rowIndex++)
            {

                dataTable.Read();
                for (int columnIndex = 0; columnIndex < dataTable.FieldCount; columnIndex++)
                {

                    // var filedType = dataTable.GetFieldType(1);
                    //判定导出类型
                    if (rowIndex == 0 && columnIndex == 1)
                    {
                        //判定导出类型
                        var text = dataTable.GetString(columnIndex);
                        switch (text)
                        {
                            case "tiny":
                                exportSchema = "tiny";
                                break;
                            case "base":
                                exportSchema = "base";
                                break;
                            default:
                                return;
                        }
                    }

                    //获取导出文件路径
                    if (rowIndex == 1 && columnIndex == 1)
                    {
                        exportPath = dataTable.GetString(columnIndex);
                    }

                    //获取Key数量
                    if (rowIndex == 2 && columnIndex == 1)
                    {
                        var countText = dataTable[columnIndex].ToString();
                        keyCount = Convert.ToInt32(countText);
                    }

                    //获取导出文件头
                    if (rowIndex == 0 && columnIndex == 4)
                    {
                        exportHeader = dataTable.GetString(columnIndex);
                    }

                    //获取导出文件尾
                    if (rowIndex == 1 && columnIndex == 4)
                    {
                        exportEnd = dataTable.GetString(columnIndex);
                    }

                    //获取列的导出参数
                    if (rowIndex == 5 && columnIndex >= 1)
                    {
                        fieldDatas[columnIndex].exportPlatform = dataTable.GetString(columnIndex);
                    }

                    //获取列的字段名
                    if (rowIndex == 6 && columnIndex >= 1)
                    {
                        fieldDatas[columnIndex].fieldName = dataTable.GetString(columnIndex);
                    }

                    //第8行才是数据列的真正开始
                    if (rowIndex < 7)
                    {
                        continue;
                    }

                    //数据行开始
                    if (rowIndex >= 7 && columnIndex >= 1)
                    {
                        //初始化FieldData
                        if (rowIndex == 7)
                        {
                            var type = dataTable.GetFieldType(columnIndex);
                            if (type != null)
                            {
                                fieldDatas[columnIndex].fieldType = type;
                                
                            }

                        }


                        var data = dataTable.GetValue(columnIndex);
                        if (data != null)
                        {
                            fieldDatas[columnIndex].dataList.Add(data.ToString());
                            fieldDatas[columnIndex].rowCount++;
                        }
                        else
                        {
                            fieldDatas[columnIndex].dataList.Add("");
                            fieldDatas[columnIndex].rowCount++;
                        }

                    }

                }

            }
            ProcessArray(exportSchema, exportPath, keyCount, exportHeader, exportEnd, fieldDatas);
        }

        private void ProcessArray(string exprotSchema, string exportPath, int keyCount, string head, string end, FieldData[] fieldDatas)
        {
            //第一行数据不参与导出
            if (fieldDatas.Length >= 1)
            {

                var fieldData = fieldDatas[1];

                var stream = new FileStream(_ExportBastPath +"\\"+  exportPath, FileMode.OpenOrCreate);
                var writer = new StreamWriter(stream);

                if (keyCount == 0)
                {
                    writer.WriteLine("[");
                }
                else
                {
                    writer.WriteLine("{");
                }

                for (int rowIndex = 0; rowIndex < fieldData.rowCount; rowIndex++)
                {

                    //处理KeyCount
                    for (int i = 1; i <= keyCount; i++)
                    {
                        var keyText = fieldDatas[i].dataList[rowIndex];
                        if (typeof(string) == fieldDatas[i].fieldType)
                        {
                            writer.WriteLine(keyText + ":" + "{");
                        }
                        else
                        {
                            writer.WriteLine('"' + keyText + '"' + ":" + "{");
                        }

                    }

                    for (int i = 1; i < fieldDatas.Length; i++)
                    {
                        var data = fieldDatas[i];

                        var type = data.fieldType;
                        if (typeof(double) == type)
                        {
                            writer.WriteLine('"' + data.fieldName + '"' + ":" + data.dataList[rowIndex] + ",");
                        }
                        else if (typeof(string) == type)
                        {
                            writer.WriteLine('"' + data.fieldName + '"' + ":" + data.dataList[rowIndex] + ",");
                        }
                    }


                    for (int i = 1; i <= keyCount; i++)
                    {
                        writer.WriteLine("}");
                    }

                    if (rowIndex != fieldData.rowCount - 1)
                    {
                        writer.WriteLine(",");
                    }

                }

                if (keyCount == 0)
                {
                    writer.WriteLine("]");
                }
                else
                {
                    writer.WriteLine("}");
                }


                writer.Flush();
                writer.Close();
                stream.Close();


            }
        }
    }
}
