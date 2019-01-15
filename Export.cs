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
        private string _KeyValueSplitChar = "= "; //json 是 :  lua 是=
        private string _ArrayFlagCharLeft = "{"; //json 是[ ,lua 是 }
        private string _ArrayFlagCharRight = "}";
        private string _WrapKey = "";
        private string _WrapString = ""; 



        public Export(string excelFileName, string exportBasePath, string exportMode, string exportType, TextBox logBox)
        {
            _ExportExcelFileName = excelFileName;
            _ExportBastPath = exportBasePath;
            _ExportMode = exportMode;
            _ExportType = exportType;
            _LogBox = logBox;
            SetExportType(exportType);
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

        private void SetExportType(string type)
        {
            if (type == "lua")
            {
                _KeyValueSplitChar = " = ";
                _ArrayFlagCharLeft = "{";
                _ArrayFlagCharRight = "}";
                _WrapKey = "";
                _WrapString = "\"";
            } else if (type == "json")
            {
                _KeyValueSplitChar = " : ";
                _ArrayFlagCharLeft = "[";
                _ArrayFlagCharRight = "]";
                _WrapKey = "\"";
                _WrapString = "\"";
            }
            else
            {
                throw new Exception("不支持的导出语言格式");
            }
        }

        private string WrapKey(string key)
        {
            return _WrapKey + key + _WrapKey;
        }

        private string  WrapString(string str , Type fieldType)
        {
            if (typeof(string) == fieldType)
            {
                return _WrapString + str + _WrapString;
            } else
            {
                return str;
            }
            
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


        private void CheckCreateDir(string dir)
        {
            var targetDir = Directory.GetParent(dir).ToString();
            if (!Directory.Exists(targetDir))
            {
                AddLog(string.Format("目录:{0} , 不存在 进行创建...", targetDir));
                Directory.CreateDirectory(targetDir);
                AddLog(string.Format("目录:{0} , 创建 完成。", targetDir));
            }
        }

        //添加导出文件头
        private void AppendHeader (StreamWriter stream , string str , int keyCount )
        {
            if (_ExportType == "lua")
            {
                stream.WriteLine(str);
            } else if (_ExportType == "json")
            {
                if (keyCount == 0)
                {
                    stream.WriteLine(_ArrayFlagCharLeft);
                }
                else
                {
                    stream.WriteLine("{");
                }
            }
        }

        //添加导出文件尾
        private void AppendEnd (StreamWriter stream , string str , int keyCount)
        {
            if (_ExportType == "lua")
            {
                stream.WriteLine(str);
            } else if (_ExportType == "json")
            {

                if (keyCount == 0)
                {
                    stream.WriteLine("\r\n" +_ArrayFlagCharRight);
                }
                else
                {
                    stream.WriteLine("\r\n" + "}");
                }
            }
        }

        private void ExportTiny(FieldData[] fieldDatas ,StreamWriter writer)
        {

            var fieldDataKey = fieldDatas[1];
            var fieldDataValue = fieldDatas[2];
            
            if (typeof(string) == fieldDataValue.fieldType)
            {
                for (int rowIndex = 0; rowIndex < fieldDataKey.rowCount; rowIndex++)
                {           
                    writer.WriteLine("\t" + WrapKey(fieldDataKey.dataList[rowIndex]) + _KeyValueSplitChar + "\"" + fieldDataValue.dataList[rowIndex] + "\",");
                }
            } else
            {
                for (int rowIndex = 0; rowIndex < fieldDataKey.rowCount; rowIndex++)
                {
                    writer.WriteLine("\t" + WrapKey(fieldDataKey.dataList[rowIndex]) + _KeyValueSplitChar + fieldDataValue.dataList[rowIndex] + ",");
                }
            }


        }

        private void ExportBase(FieldData[] fieldDatas, StreamWriter writer,int keyCount)
        {
            var fieldData = fieldDatas[1];
            for (int rowIndex = 0; rowIndex < fieldData.rowCount; rowIndex++)
            {

                var tabPreFix = "";
                //处理KeyCount
                for (int i = 1; i <= keyCount; i++)
                {
                    var keyText = fieldDatas[i].dataList[rowIndex];
                    if (typeof(string) == fieldDatas[i].fieldType)
                    {
                        writer.Write(keyText + _KeyValueSplitChar + "{");
                    }
                    else
                    {
                        writer.Write(WrapKey(keyText) + _KeyValueSplitChar + "{");
                    }

                    tabPreFix += "\t";

                }

                writer.Write("\r\n");

                //导出数据部分
                for (int i = 1; i < fieldDatas.Length; i++)
                {
                    var data = fieldDatas[i];

                    var type = data.fieldType;
                        
                    if (typeof(double) == type || typeof(string) == type)
                    {

                        if (i != 1)
                        {
                            writer.Write(",\r\n");
                        }

                        writer.Write(tabPreFix + WrapKey(data.fieldName) + _KeyValueSplitChar + WrapString(data.dataList[rowIndex], type));


                    }
            
                    
                }

                for (int i = 1; i <= keyCount; i++)
                {
                    writer.Write("}");
                }

                if (rowIndex != fieldData.rowCount - 1)
                {
                    writer.WriteLine(",");
                }

            }

        }

        private void ProcessArray(string exprotSchema, string exportPath, int keyCount, string head, string end, FieldData[] fieldDatas)
        {
            //如果只有一列数据 那么不导出 因为第一列是备注说明
            if (fieldDatas.Length <= 1)
            {
                AddLog(string.Format("欲导出文件: {0 },数据列数为 :0 , 跳过 ", exportPath));
                return;
            }


            var fieldData = fieldDatas[1];
            var targetPath = _ExportBastPath + "\\" + exportPath;
            CheckCreateDir(targetPath);
               

            var stream = new FileStream(_ExportBastPath +"\\"+  exportPath, FileMode.Create);
            var writer = new StreamWriter(stream);

            //如果是tiny 并且是导出为json 那么这里恒为1 否则会被判定为array.
            if (exprotSchema == "tiny" && _ExportType == "json")
            {
                keyCount = 1;
            }

            AppendHeader(writer, head  ,keyCount);
              

            if (exprotSchema == "tiny")
            {
                ExportTiny(fieldDatas, writer);
            }else if(exprotSchema == "base")
            {
                ExportBase(fieldDatas, writer, keyCount);
            }


            AppendEnd(writer, end , keyCount);

            writer.Flush();
            writer.Close();
            stream.Close();

        }
    }
}
