using System;
using System.Collections.Generic;

public class AraryFieldData
{
    public string keyName;
    public List<string> values;
}

public class ItemData
{
    public int type;
    public int id;
    public int count;
}

public enum FieldObjectType
{
    BOOLEAN = 1,
    STRING = 2,
    NUMBER = 3,
    ITEM = 4,
    ARRAY = 5
}

public class FieldData
{
    public string exportPlatform;
    public string fieldName;
    public Type fieldType;
    public List<string> dataList = new List<string>(); //普通字段

    public FieldObjectType objType;
    public List<Object> objList = new List<Object>(); //对象列表根据 dataList 序列化后的字段
    
    public void AddArrayField(string keyName, FieldData data)
    {
        if (objList == null)
        {
            objList = new List<Object>();
        }

        AraryFieldData adt = new AraryFieldData();
        adt.keyName = keyName;
        adt.values = data.dataList;

        objList.Add(adt);
    }

    public int ObjListCount()
    {
        return objList.Count;
    }

    public AraryFieldData GetArrayFieldByIndex(int index)
    {
        if (objType != FieldObjectType.ARRAY)
        {
            return null;
        }

       if  (index >= 0 && index < objList.Count)
        {
            return objList[index] as AraryFieldData;
        }

        return null;
    }

    public List<ItemData> GetItemDataListByIndex(int index)
    {
        if (objType != FieldObjectType.ITEM)
        {
            return null;
        }

        if (index >= 0 && index < objList.Count)
        {
            return objList[index] as List<ItemData>;
        }

        return null;
    }

    public void AddItemList(List<ItemData> itemList)
    {
        if (objList == null)
        {
            objList = new List<Object>();
        }

        objList.Add(itemList);

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
