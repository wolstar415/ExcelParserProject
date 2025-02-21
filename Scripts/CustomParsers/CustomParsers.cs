using System;
using UnityEngine;

#region Custom Parser Interfaces and Implementations


public readonly struct SetParser : IMultiColumnParser
{
    // values[0] -> setId, values[1] -> setLevel
    public object Parse(params string[] values)
    {
        if (values.Length < 2)
            throw new Exception($"Not enough columns to parse Set object. Got {values.Length} columns.");

        var set = new Set();
        set.Name = values[0];             // string 값 그대로 사용
        set.Value = values[1]; // int 값 파싱
        Debug.Log($"Name: {set.Name} value : {set.Value} ");
        return set;
    }
}

public class Set
{
    public string Name;
    public string Value;
}


public readonly struct WeightedValue<T> : ICustomParser
{
    public readonly T value;
    public readonly float weight;

    public WeightedValue(T value, float weight)
    {
        this.value = value;
        this.weight = weight;
    }

    public override string ToString() => $"{value} ({weight})";

    public object Parse(string value)
    {
        return ParseValue(value);
    }

    // 예: "Apple:0.75" → value="Apple", weight=0.75

    public static WeightedValue<T> ParseValue(string value)
    {
        var parts = value.Split(':');
        if (parts.Length < 2)
            throw new Exception("Invalid format for WeightedValue. Expected: value:weight");
        T val = (T)Convert.ChangeType(parts[0].Trim(), typeof(T));
        float w = float.Parse(parts[1].Trim());
        return new WeightedValue<T>(val, w);
    }
}

#endregion
