using System;
using System.Collections.Generic;

public class AraryFieldData
{
    public string keyName;
    public List<string> values;
}

public class FieldData
{
    public string exportPlatform;
    public string fieldName;
    public Type fieldType;
    public List<string> dataList = new List<string>(); //普通字段
    public List<AraryFieldData> arrayList; //整理后的数组字段
    
    public void AddArrayField(string keyName, FieldData data)
    {
        if (arrayList == null)
        {
            arrayList = new List<AraryFieldData>();
        }

        AraryFieldData adt = new AraryFieldData();
        adt.keyName = keyName;
        adt.values = data.dataList;

        arrayList.Add(adt);
    }

    public bool IsArrayField()
    {
        return arrayList != null;
    }


    public void Add(string str)
    {
        dataList.Add(str);
    }

    public int RowCount
    {
        get { return dataList.Count; }
    }

    public string GetString(int rowIndex)
    {
        if (rowIndex >= 0 && rowIndex < dataList.Count)
        {
            return dataList[rowIndex];
        } else
        {
            return "";
        }
          
    }

    public bool CanExportTo(string plat)
    {
        return exportPlatform.IndexOf(plat) >= 0;

    }

	public FieldData()
	{

	}

    public void CheckRealRowCount()
    {
        for (var i = dataList.Count - 1; i >= 7 ; i--)
        {
            if (dataList[i] == "")
            {
                dataList.RemoveAt(i);
            }
        }
    }
}
