using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Exporter
{
    public class ExporterJson : ExporterBase
    {
        public ExporterJson(string exportMode) : base(exportMode)
        {
        }
        protected override void AddHeader(ExcelSheetData data, StreamWriter writer)
        {
            writer.WriteLine("[");
        }

        protected override void AddEnd(ExcelSheetData data, StreamWriter writer)
        {
            writer.WriteLine("]");
        }

        protected override void ExportTiny(ExcelSheetData data, StreamWriter writer)
        {
            var fieldDataKey = data.filedList[0];
            var fieldDataValue = data.filedList[1];

            writer.WriteLine("{");

            for (var rowCount = 0; rowCount < fieldDataKey.dataList.Count; rowCount++)
            {
                string str = "";
                int intRes = -1;
                double doubleRes = 0;
                string value = fieldDataValue.dataList[rowCount];
                if (int.TryParse(value, out intRes))
                {
                    str = string.Format("\"{0}\":{1}", fieldDataKey.dataList[rowCount], intRes);
                }
                else if (double.TryParse(value, out doubleRes))
                {
                    str = string.Format("\"{0}\":{1}", fieldDataKey.dataList[rowCount], doubleRes);
                }
                else
                {
                    if (value.StartsWith("[") && value.EndsWith("]"))
                    {
                        str = string.Format("\"{0}\":{1}", fieldDataKey.dataList[rowCount], value);
                    }
                    else
                    {
                        str = string.Format("\"{0}\":\"{1}\"", fieldDataKey.dataList[rowCount], value);
                    }
                }

                if (rowCount != fieldDataKey.dataList.Count - 1)
                {
                    str += ",";
                }
                writer.WriteLine(str);
            }
            writer.WriteLine("}");
        }

        protected override void ExportBase(ExcelSheetData data, StreamWriter writer, string exportMode)
        {
            FieldData fieldData = null;
            var rowCount = data.filedList[0].RowCount;
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {

                //处理keyCount
                string key = "";
                for (var i = 0; i < data.keyCount; i++)
                {
                    //key += string.Format("\"{0}\" : ", data.filedList[i].dataList[rowIndex]) + "{";
                    key += "{";
                }

                string appendStr = (key + "\n");

                for (var i = 0; i < data.filedList.Count; i++)
                {
                    fieldData = data.filedList[i];

                    if (!fieldData.CanExportTo(exportMode))
                    {
                        continue;
                    }

                    //数组列表
                    if (fieldData.IsArrayField())
                    {
                        string str = "   " + fieldData.fieldName + "[";
                        for (var arrayIndex = 0; arrayIndex < fieldData.arrayList.Count; arrayIndex++)
                        {
                            AraryFieldData afd = fieldData.arrayList[arrayIndex];
                            str = str + "{" + string.Format("\"{0}\" : {1}", afd.keyName, afd.values[rowIndex]) + "}";

                            if (arrayIndex != fieldData.arrayList.Count - 1)
                            {
                                str = str + ",";
                            }
                        }

                        str += "]";

                        appendStr += str;
                    }
                    else
                    {
                        var tempData = fieldData.dataList[rowIndex];
                        //普通字段
                        if (fieldData.fieldType == typeof(string))
                        {
                            string str;
                            if (tempData.StartsWith("{") && tempData.EndsWith("}"))
                            {
                                str = string.Format(" \"{0}\":{1}", fieldData.fieldName, tempData);
                            }
                            else
                            {
                                str = string.Format(" \"{0}\":\"{1}\"", fieldData.fieldName, tempData);
                            }
                            appendStr += str;
                        }
                        else
                        {
                            string str = string.Format(" \"{0}\":{1}", fieldData.fieldName, tempData);
                            appendStr += str;
                        }

                    }

                    appendStr += ",\n";
                }


                appendStr = appendStr.Substring(0, appendStr.Length - 2);
                writer.WriteLine(appendStr);

                if (appendStr.Length > 2)
                {
                    if (rowIndex == rowCount - 1)
                    {
                        writer.WriteLine("  }");
                    }
                    else
                    {
                        writer.WriteLine("  },");
                    }
                }

            }

        }
    }
}
