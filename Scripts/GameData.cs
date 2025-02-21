using UnityEngine;          // Unity에서 Debug.Log 사용
using System;               // Convert 등
using System.Collections.Generic;

using System.Numerics;      // Vector2, Vector3 등

// ★ GameData: 모든 파싱 결과가 들어가는 컨테이너
public class GameData
{
    // UnitData
    public Dictionary<string, UnitData> UnitData = new ();
}

public enum ABCD
{
A1,
B2,

}

// ★ UnitData (필드만 사용)
public class UnitData
{
    public string id;
    public string name;
    public ABCD abc;
    // 쉼표 구분 "10,20" → [10,20]
    public int[] bonusDamage;

    // 커스텀 파서 (Vector2)
    [CustomParser(typeof(Vector2Parser))]
    public UnityEngine.Vector2 position;

    // Key() 예시 (없으면 Fallback)
    public string Key() => $"{id}";

}

/// <summary>
/// Enum 예시
/// </summary>
public enum UnitType
{
    Melee,
    Ranged,
    Magic
}



// ★ ItemData (필드만 사용)
public class ItemData
{
    public string itemId;
    public string itemName;

    // 쉼표 구분 "Potion,Heal" → ["Potion","Heal"]
    public string[] tags;

    // 쉼표 구분 "10,20,30" → List<int> {10,20,30}
    public List<int> stats;

    // Key() 없음 → Fallback 사용 or Dict에 안 넣을 수도
}

// ★ AnotherData (멀티 컬럼 파서 예시)
public class AnotherData
{
    public int code;

    [MultiColumnParser(typeof(SetParser), "valStr", "valInt")]
    public Set combined;

    // Key() 없음 → Fallback
}

// ★ 예시 Set 구조
public class Set
{
    public string StrVal;
    public int IntVal;
}

//
// ─────────────────────────────────────────────────────────────────────────────────
//   2) Parser Interfaces & Attributes (필드 대상)
// ─────────────────────────────────────────────────────────────────────────────────
//

public interface ICustomParser
{
    object Parse(string value);
}

public interface IMultiColumnParser
{
    object Parse(params string[] values);
}

// 필드에도 적용 가능하도록 설정
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class CustomParserAttribute : Attribute
{
    public Type ParserType { get; }
    public CustomParserAttribute(Type parserType)
    {
        ParserType = parserType;
    }
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class MultiColumnParserAttribute : Attribute
{
    public Type ParserType { get; }
    public string[] ColumnNames { get; }
    public MultiColumnParserAttribute(Type parserType, params string[] columnNames)
    {
        ParserType = parserType;
        ColumnNames = columnNames;
    }
}

//
// ─────────────────────────────────────────────────────────────────────────────────
//   3) Parser Implementations (Vector2, SetParser)
// ─────────────────────────────────────────────────────────────────────────────────
//

public class Vector2Parser : ICustomParser
{
    public object Parse(string value)
    {
        var parts = value.Split(',');
        if (parts.Length != 2)
            throw new Exception($"Invalid Vector2 format: {value}");
        float x = float.Parse(parts[0].Trim());
        float y = float.Parse(parts[1].Trim());
        return new UnityEngine.Vector2(x, y);
    }
}

public class SetParser : IMultiColumnParser
{
    public object Parse(params string[] values)
    {
        if (values.Length < 2)
            throw new Exception($"Not enough columns for SetParser. Need 2, got {values.Length}");
        var setObj = new Set();
        setObj.StrVal = values[0];
        setObj.IntVal = int.Parse(values[1]);
        return setObj;
    }
}
