using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Exporter
{
    public class ExporterLua : ExporterBase
    {
        public ExporterLua(string exportMode) : base(exportMode)
        {
        }

        protected override void ExportTiny(ExcelSheetData data, StreamWriter writer)
        {
            var fieldDataKey = data.filedList[0];
            var fieldDataValue = data.filedList[1];

            writer.WriteLine("{");

            for (var rowCount = 0; rowCount < fieldDataKey.dataList.Count; rowCount++)
            {
                string str = "";
                str = string.Format("{0}=\"{1}\"", fieldDataKey.dataList[rowCount], fieldDataValue.dataList[rowCount]);

                if (rowCount != fieldDataKey.dataList.Count)
                {
                    str += ",";
                }
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
                    key += string.Format("[{0}] = ", data.filedList[i].dataList[rowIndex]) + "{";
                }

                string appendStr = key + "\n";

                for (var i = 0; i < data.filedList.Count; i++)
                {
                    fieldData = data.filedList[i];

                    if (!fieldData.CanExportTo("s"))
                    {
                        continue;
                    }

                    //数组列表
                    if (fieldData.IsArrayField())
                    {
                        string str = "    " + fieldData.fieldName + " = {";
                        for (var arrayIndex = 0; arrayIndex < fieldData.arrayList.Count; arrayIndex++)
                        {
                            AraryFieldData afd = fieldData.arrayList[arrayIndex];
                            str = str + string.Format("[{0}]={1}", afd.keyName, afd.values[rowIndex]);

                            if (arrayIndex != fieldData.arrayList.Count - 1)
                            {
                                str = str + "}";
                            }
                        }

                        appendStr += str;
                    }
                    else
                    {
                        //普通字段
                        if (fieldData.fieldType == typeof(string))
                        {
                            string str = string.Format("    {0}=\"{1}\"", fieldData.fieldName, fieldData.dataList[rowIndex]);
                            appendStr += str;
                        }
                        else
                        {
                            string str = string.Format("    {0}={1}", fieldData.fieldName, fieldData.dataList[rowIndex]);
                            appendStr += str;
                        }
                    }

                    appendStr += ",\n";
                }

                appendStr = appendStr.Substring(0, appendStr.Length - 2);
                writer.WriteLine(appendStr);

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
