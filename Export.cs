using System;
using ExcelDataReader;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;

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
        private int _nExportType = 1;
        private List<String> _rowComent = new List<String>();

        private const int JSON = 1;
        private const int LUA = 2;

        private bool _isExportCode = false;

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
                _nExportType = LUA;
            } else if (type == "json")
            {
                _KeyValueSplitChar = " : ";
                _ArrayFlagCharLeft = "[";
                _ArrayFlagCharRight = "]";
                _WrapKey = "\"";
                _WrapString = "\"";
                _nExportType = JSON;
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
            } else if (typeof(Boolean) == fieldType)
            {

                if (str.Length <= 0)
                {
                    return "false";
                }

              if (str[0] == 'F' || str[0] == 'f')
                {
                    return "false";
                } else
                {
                    return "true";
                }
            }
            else
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
            _isExportCode = false;
            //建立所有会缓存的数组
            var fieldDatas = new FieldData[dataTable.FieldCount];
            for (int i = 0; i < dataTable.FieldCount; i++)
            {
                fieldDatas[i] = new FieldData();
            }


            _rowComent.Clear();
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
                        if (dataTable.GetFieldType(columnIndex) != typeof(string))
                        {
                            return;
                        }
                        var text = dataTable.GetString(columnIndex);
                        switch (text)
                        {
                            case "tiny":
                                exportSchema = "tiny";
                                break;
                            case "base":
                                exportSchema = "base";
                                break;
                            case "codetiny":
                                exportSchema = "tiny";
                                _isExportCode = true;
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

                    //读取备注说明
                    if (columnIndex == 0)
                    {
                        var data = dataTable.GetValue(columnIndex);
                        if (data != null)
                        {
                            _rowComent.Add(data.ToString());
         
                        }
                        else
                        {
                            _rowComent.Add("");
                        }
                    }

                }



            }

            //去掉空行
            var FieldDataList = new List<FieldData>();
            for (var i = 0; i < fieldDatas.Length; i++)
            {
                if ( i == 0 || fieldDatas[i].fieldName != null)
                    FieldDataList.Add(fieldDatas[i]);
            }

            //复制回去
            fieldDatas = new FieldData[FieldDataList.Count];
            for (var i = 0; i < FieldDataList.Count; i++)
            {
                fieldDatas[i] = FieldDataList[i];
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
                    writer.Write("\t" + WrapKey(fieldDataKey.dataList[rowIndex]) + _KeyValueSplitChar + "\"" + fieldDataValue.dataList[rowIndex] + "\",");

             
                    if (_nExportType == LUA && _isExportCode && _rowComent[rowIndex] != "")
                    {
                        writer.Write("\t" + "--" + _rowComent[rowIndex]);
                    } else
                    {
                        writer.Write("\r\n");
                    }
                }
            } else
            {
                for (int rowIndex = 0; rowIndex < fieldDataKey.rowCount; rowIndex++)
                {
                    writer.Write("\t" + WrapKey(fieldDataKey.dataList[rowIndex]) + _KeyValueSplitChar + fieldDataValue.dataList[rowIndex] + ",");


                    if (_nExportType == LUA && _isExportCode && _rowComent[rowIndex] != "")
                    {
                        writer.Write("\t" + "--" + _rowComent[rowIndex] + "\r\n");
                    }
                    else
                    {
                        writer.Write("\r\n");
                    }
                }
            }
          

        }

        private void ExportBase(FieldData[] fieldDatas, StreamWriter writer,int keyCount)
        {
            var fieldData = fieldDatas[1]; //拿第一列数据的长度作为循环遍历次数 实际上这个有点问题 如果第一行长度不是最长的 后面的数据都会导不出来。 暂时先这样实现 

            for (int rowIndex = 0; rowIndex < fieldData.rowCount; rowIndex++)
            {

                var tabPreFix = "";
                //处理KeyCount
                for (int i = 1; i <= keyCount; i++)
                {
                    var keyText = fieldDatas[i].dataList[rowIndex];
                    if (_nExportType == LUA)
                    {
                        if (typeof(string) == fieldDatas[i].fieldType)
                        {
                            writer.Write(keyText + _KeyValueSplitChar + "{");
                        }
                        else
                        {
                            writer.Write("[" + keyText + "]" + _KeyValueSplitChar + "{");
                        }

                    }
                    else {
                        if (typeof(string) == fieldDatas[i].fieldType)
                        {
                            writer.Write(keyText + _KeyValueSplitChar + "{");
                        }
                        else
                        {
                            writer.Write(WrapKey(keyText) + _KeyValueSplitChar + "{");
                        }
                    }
                    tabPreFix += "\t";

                }

                writer.Write("\r\n");

                //导出数据部分
                Dictionary<string, string> arrayMap = new Dictionary<string, string>();


                bool skipDouhao = false;
                for (int i = 1; i < fieldDatas.Length; i++)
                {
                    var data = fieldDatas[i];

                    var type = data.fieldType;

                    if (null == type)
                    {
                        continue;
                    }

                   if (!data.CanExportTo(_ExportMode))
                    {
                        continue;
                    }

                    if (typeof(double) == type || typeof(string) == type || typeof(Boolean) == type)
                    {

                        //第一行不插入 换行逗号 如果是数组缓存 字段那么也跳过。
                        if (i != 1)
                        {
                            if (!skipDouhao)    
                            {
                                writer.Write(",\r\n");
                            }
                           
                        }

                        if (data.fieldName.StartsWith("_"))
                        {
                            var pos = data.fieldName.IndexOf(":");
                            if (pos >= 0)
                            {
                                var arrayTableName = data.fieldName.Substring(1, pos - 1);
                                var index = data.fieldName.Substring(pos + 1, data.fieldName.Length - pos - 1);
                                string text = "";
                                if (!arrayMap.TryGetValue(arrayTableName, out text))
                                {
                                    text = "{ ";
                                }

                                text = text + "[" + WrapKey(index) + "]" + _KeyValueSplitChar + WrapString(data.dataList[rowIndex], type) + " , ";

                                arrayMap[arrayTableName] = text;

                                skipDouhao = true; //走了这里说明这个表字段被临时缓存起来了 下一行啥也不用干。


                            } else
                            {
                                throw new Exception("找到 _ 但是没有找到 : ,字段名:" + data.fieldName);
                            }

                        } else if (data.fieldName.StartsWith("#A")) {
                            var fieldName = data.fieldName.Substring(3, data.fieldName.Length - 3);
                            var tb = data.dataList[rowIndex].Split('|');
                            string exportStr = "";
                            foreach(var str in tb)
                            {
                                var itemValue = str.Split(',');
                                if (itemValue.Length >= 2 ){
                                    string typeofItemIDStr = itemValue[0];
                                    string countStr = itemValue[1];

                                    int id = Convert.ToInt32(typeofItemIDStr);
                                    int count = Convert.ToInt32(countStr);

                                    if (_nExportType == LUA)
                                    {
                                        exportStr = exportStr + "{ id = " + id.ToString() + ",count=" + count.ToString() + "},";
                                    } else
                                    {
                                        exportStr = exportStr + "{ \"id:\"" + id.ToString() + ",\"count\":" + count.ToString() + "},";
                                    }
                                   
                                }
                            }

                            if (_nExportType == LUA)
                            {
                                exportStr = "{" + exportStr.Substring(0, exportStr.Length - 1) + "}";
                            } else
                            {
                                exportStr = "[" + exportStr.Substring(0, exportStr.Length - 1) + "]";
                            }

                            writer.Write(tabPreFix + WrapKey(fieldName) + _KeyValueSplitChar + exportStr);
                            skipDouhao = false;

                        } else {

                            skipDouhao = true;
                            //对应类型如果是默认值全部跳过生成对应的字段 节省内存。
                            if (data.dataList[rowIndex] == "")
                            {
                                continue;
                            }

                            if (typeof(Boolean) == type && data.dataList[rowIndex] == "False")
                            {
                                continue;
                            }

                            if (typeof(double) == type && data.dataList[rowIndex] == "0")
                            {
                                continue;
                            }



                            writer.Write(tabPreFix + WrapKey(data.fieldName) + _KeyValueSplitChar + WrapString(data.dataList[rowIndex], type));
                            skipDouhao = false;//走了这里 下一行还是要正常插入逗号
                        }

                    }
                    else
                    {
                        throw new Exception("未处理的字段类型:" + type.ToString() + "字段名:" + data.fieldName);
                    }
                          
                }
                List<string> keys = new List<string>(arrayMap.Keys);
   
                for (var i = 0; i < keys.Count; i ++)
                {
                    var value = arrayMap[keys[i]];
                    value = value.Substring(0, value.Length - 2); //去掉尾部的两个字符 逗号 和 空格 用于排版的
                    writer.Write(tabPreFix + WrapKey(keys[i]) + _KeyValueSplitChar + value + "}");
                }


                for (int i = 1; i <= keyCount; i++)
                {
                    writer.Write("\r\n}");
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

            string fileName = _ExportBastPath + "\\" + exportPath;

            if (_isExportCode)
            {
                if (_nExportType == JSON)
                {
                    fileName += ".js";
                }
                else
                {
                    fileName += ".lua";
                }
            }

            var stream = new FileStream(fileName , FileMode.Create);
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

            if (_isExportCode)
            {
                if (_nExportType == JSON)
                {

                } else
                {
                    var pos = head.IndexOf("=");
                    string tableName = head.Substring(0, pos);
                    writer.WriteLine("return " + tableName);
                }
            }

            writer.Flush();
            writer.Close();
            stream.Close();

        }
    }
}
