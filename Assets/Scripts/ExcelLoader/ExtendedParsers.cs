using System;
using UnityEngine;

public readonly struct Vector2Parser : ICustomParser
{
    public object Parse(string value)
    {
        return ParseValue(value);
    }

    public static Vector2 ParseValue(string value)
    {
        var parts = value.Split(',');
        if (parts.Length == 0)
            throw new Exception($"Invalid Vector2 format: {value}");

        if (parts.Length == 1)
        {
            return new Vector2(float.Parse(parts[0].Trim()), 0);
        }

        return new Vector2(float.Parse(parts[0].Trim()), float.Parse(parts[1].Trim()));
    }
}


public readonly struct Vector3Parser : ICustomParser
{
    public object Parse(string value)
    {
        return ParseValue(value);
    }

    public static Vector3 ParseValue(string value)
    {
        var parts = value.Split(',');
        if (parts.Length == 0)
            throw new Exception($"Invalid Vector2 format: {value}");

        if (parts.Length == 1)
        {
            return new Vector3(float.Parse(parts[0].Trim()), 0);
        }

        if (parts.Length == 2)
        {
            return new Vector3(float.Parse(parts[0].Trim()), float.Parse(parts[1].Trim()));
        }

        return new Vector3(float.Parse(parts[0].Trim()), float.Parse(parts[1].Trim()), float.Parse(parts[2].Trim()));
    }
}