using System.Collections.Generic;

public class ExcelData
{
    public Dictionary<string, UnitData> UnitData;
}

public class UnitData
{
    public string Key() => $"{id}.{name}";
    public string id;
    public string name;
}

public class ItemData
{
    public string id;
    public string name;
}