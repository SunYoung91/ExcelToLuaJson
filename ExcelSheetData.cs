using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class ExcelSheetData
{
    public List<FieldData> filedList;
    public string sheetName;
    public string exportSchema; //导出的类型 tiny or base
    public string exportPath;
    public string exportHeader;
    public string exportEnd;
    public int keyCount;
    public bool isNeedExprot;

    public ExcelSheetData()
    {

    }

    public void ReadFromExcel(string sheetName, IExcelDataReader dataTable)
    {

        this.sheetName = sheetName;
        filedList = new List<FieldData>();


        for (int i = 0; i < dataTable.FieldCount; i++)
        {
            filedList.Add(new FieldData());
        }


        for (int rowIndex = 0; rowIndex < dataTable.RowCount; rowIndex++)
        {
            dataTable.Read();
            for (int columnIndex = 0; columnIndex < dataTable.FieldCount; columnIndex++)
            {
                FieldData fieldData = filedList[columnIndex];

                //根据有效数据的第一行决定本行的实际类型 也就是第一行数据一定不能为空
                if (rowIndex == 7)
                {
                    var type = dataTable.GetFieldType(columnIndex);
                    if (type != null)
                    {
                        fieldData.fieldType = type;
                    }
                }

                var data = dataTable.GetValue(columnIndex);
                if (data != null)
                {
                    fieldData.dataList.Add(data.ToString());
                }
                else
                {
                    fieldData.dataList.Add("");

                }

            }
        }


        var exportSchema = GetCellString(0, 1);
        isNeedExprot = true;
        if (exportSchema != "base" && exportSchema != "tiny"){
            isNeedExprot = false;
            return;
        }

        InitExportBaseInfo();

        DeleteNoneDateCell();//删掉非数据行的字段
        MergeArrayField();
    }

    private string GetCellString(int rowIndex, int columnIndex)
    {
        if (filedList.Count < columnIndex)
            return "";

        FieldData fieldData = filedList[columnIndex];

        if (fieldData.RowCount < rowIndex)
        {
            return "";
        }

        return fieldData.GetString(rowIndex);

    }

    private void InitExportBaseInfo()
    {
        exportSchema = GetCellString(0, 1);

        exportPath = GetCellString(1, 1);

        //获取Key数量
        var countText = GetCellString(2, 1);
        keyCount = Convert.ToInt32(countText);

        //获取导出文件头
        exportHeader = GetCellString(0, 4);

        //获取导出文件尾
        exportEnd = GetCellString(1, 4);


        //获取列的导出参数
        for (var i = 0; i < filedList.Count; i++)
        {
            filedList[i].exportPlatform = filedList[i].GetString(5);
        }

        //获取列的字段名
        for (var i = 0; i < filedList.Count; i++)
        {
            filedList[i].fieldName = filedList[i].GetString(6);
        }

        //过滤掉 filedName为 空的字段
        for (var i = filedList.Count - 1; i > 0; i--)
        {
            if (filedList[i].fieldName == "")
            {
                filedList.RemoveAt(i);
            }
        }

        //第一行恒为备注行 去掉
        filedList.RemoveAt(0);

        //有时候row 中间可能会空一行 这里确定row 的最终有效长度
        FieldData fieldData = filedList[0];
        fieldData.CheckRealRowCount();
    }

    private void MergeArrayField() //合并以 _开头的数组字段加index
    {
        Dictionary<string, FieldData> arrayField = new Dictionary<string, FieldData>();

        //遍历检查是不是以_ 开头
        for (var i = filedList.Count - 1; i > 0; i--)
        {
            FieldData field = filedList[i];
            if (field.fieldName.StartsWith("_"))
            {
                var pos = field.fieldName.IndexOf(":");
                if (pos >= 0)
                {
                    var arrayTableName = field.fieldName.Substring(1, pos - 1);
                    var index = field.fieldName.Substring(pos + 1, field.fieldName.Length - pos - 1);

                    FieldData arrayData = null;
                    if (!arrayField.TryGetValue(arrayTableName, out arrayData))
                    {
                        arrayData = new FieldData();
                        arrayData.fieldName = arrayTableName;
                        arrayField.Add(arrayTableName, arrayData);
                    }
                    arrayData.AddArrayField(index, field);

                    filedList.RemoveAt(i); //移除此列
                }
            }
        }

        //遍历已经合并的临时数组列 加到 filedList 中
        foreach (KeyValuePair<string, FieldData> kv in arrayField)
        {
            filedList.Add(kv.Value);
        }

    }

    //删掉非数据行
    private void DeleteNoneDateCell() 
    {
        for (var i = 0; i < filedList.Count; i++)
        {
            FieldData field = filedList[i];
            field.dataList.RemoveRange(0, 7);     
        }
    }
}
