using System;
using System.Collections.Generic;

public class FieldData
{
    public string exportPlatform;
    public string fieldName;
    public Type fieldType;
    public List<string> dataList = new List<string>() ;
    public int rowCount = 0;
    public bool CanExportTo(string plat)
    {
        return exportPlatform.IndexOf(plat) >= 0;

    }
	public FieldData()
	{

	}
}
