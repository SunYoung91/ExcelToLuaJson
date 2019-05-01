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
        protected override void ExportTiny(ExcelSheetData data, StreamWriter writer)
        {
            var fieldDataKey = data.filedList[0];
            var fieldDataValue = data.filedList[1];
            FieldData typeField = null;

            if (data.filedList.Count > 2)
            {
                typeField = data.filedList[2];
            }

            for (var rowCount = 0; rowCount < fieldDataKey.dataList.Count; rowCount++)
            {

                var ttype = "";
                if ((typeField != null) && (typeField.dataList[rowCount] != ""))
                {
                    ttype = typeField.dataList[rowCount];
                }


                string str = "";
                switch (ttype)
                {
                    case "int":
                        {
                            str = string.Format("{0}={1}", fieldDataKey.dataList[rowCount], fieldDataValue.dataList[rowCount]);
                        }
                        break;
                    case "int_array":
                        {
                            var values = fieldDataValue.dataList[rowCount].Split(',');
                            var arrayValues = "{";
                            foreach (var v in values)
                            {
                                arrayValues += v + ",";
                            }

                            arrayValues = arrayValues.Substring(0, arrayValues.Length - 1) + "}";

                            str = string.Format("{0}={1}", fieldDataKey.dataList[rowCount], arrayValues);
                        }
                        break;
                    default: //默认当做string 处理
                        {
                            str = string.Format("{0}:\"{1}\"", fieldDataKey.dataList[rowCount], fieldDataValue.dataList[rowCount]);
                        }
                        break;

                }

                if (rowCount != fieldDataKey.dataList.Count - 1)
                {
                    str += ",";
                }

                writer.WriteLine(str);
            }
        }

        protected override void ExportBase(ExcelSheetData data, StreamWriter writer)
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

                writer.WriteLine(key);

                for (var i = 0; i < data.filedList.Count; i++)
                {
                    fieldData = data.filedList[i];           
                    //数组列表
                    if (fieldData.objType == FieldObjectType.ARRAY)
                    {
                        string str = "    " + fieldData.fieldName + "={";
                        for (var arrayIndex = 0; arrayIndex < fieldData.ObjListCount(); arrayIndex++)
                        {
                            AraryFieldData afd = fieldData.GetArrayFieldByIndex(arrayIndex);
                            str = str + string.Format("[{0}]={1}", afd.keyName, afd.values[rowIndex]);

                            if (arrayIndex != fieldData.ObjListCount() - 1)
                            {
                                str = str + ",";
                            }
                        }


                        writer.Write(str + "}");

                    } else if (fieldData.objType == FieldObjectType.ITEM)
                    {
                        string str = "   " + fieldData.fieldName +  " = {";
                        
                        List<ItemData> itemList = fieldData.GetItemDataListByIndex(rowIndex);

                        for (var itemIndex = 0; itemIndex < itemList.Count; itemIndex++)
                        {
                            ItemData item = itemList[itemIndex];

                            str = str + "{" + string.Format("type={0},id={1},count={2}", item.type, item.id, item.count) + "}";

                            if (itemIndex != itemList.Count - 1)
                            {
                                str = str + ",";
                            }
                        }

         
                        str += "}";

                        writer.Write(str);

                    }
                    else
                    {
                        //普通字段
                        if (fieldData.fieldType == typeof(string))
                        {
                            string str = string.Format("    {0}=\"{1}\"", fieldData.fieldName, fieldData.dataList[rowIndex]);
                            writer.Write(str);
                        }
                        else if (fieldData.fieldType == typeof(Boolean))
                        {
                            //读取出来的true 和 false 都是大写开头的这里要单独处理一下
                            string str = string.Format("    {0}={1}", fieldData.fieldName, fieldData.dataList[rowIndex].ToLower());
                            writer.Write(str);
                        }
                        else
                        {
                            string str = string.Format("    {0}={1}", fieldData.fieldName, fieldData.dataList[rowIndex]);
                            writer.Write(str);
                        }
                    }

                    //最后一行不给逗号分隔
                    if (i != data.filedList.Count - 1)
                    {
                        writer.WriteLine(",");
                    }
                    else
                    {
                        writer.WriteLine(" ");
                    }                  
                }
                if(rowIndex == rowCount - 1)
                {
                    writer.WriteLine("  }");
                } else
                {
                    writer.WriteLine("  },");
                }
               
            }

        }

    }

}
