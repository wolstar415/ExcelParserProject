using System;
using System.Collections.Generic;

[Serializable]
public class ExcelData
{
    public List<MonsterData> MonsterData;

    [SheetBinding(sheetName: "MonsterData")]
    //Key => first column : id
    public Dictionary<string,MonsterData> MonsterDic;


    public PcData PcData;

    [SheetBinding(sheetName: "PcData")]
    public List<PcData> PcDataList;


    [SheetBinding(sheetName: "PcData")]
    //Key => Key();
    public Dictionary<string,PcData> PcDataDic;
}

[Serializable]
public class MonsterData
{
    public string id;
    public string name;
    public float hp;
    public float attack;
    public int exp;
    public AttackType attackType;
}

[Serializable]
public class PcData
{
    public string Key() => $"Pc_{id}";
    public string id;
    public string name;
    public float[] attack;
    public float hp;
    
    [ExcelParer(columnName:"pcClass")]
    public PCClass type;

    public int exp;
}


public enum PCClass
{
    Knight,
    Mage,
    Archer,
    Healer,
    Sorceress,
    Rogue,
    Warrior,
}
public enum AttackType
{
    melee,
    ranged,
}