using System;

public interface ICustomParser
{
    object Parse(string value);
}

/// <summary>
/// IMultiColumnParser 인터페이스: 여러 컬럼을 받아 하나의 값으로 파싱합니다.
/// </summary>
public interface IMultiColumnParser
{
    object Parse(params string[] values);
}



[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class ExcelParerAttribute : Attribute
{
    /// <summary>연결할 Column</summary>
    public string ColumnName { get; set; }
    /// <summary> 무시하는 데이터</summary>
    public bool Ignore { get; set; }
    /// <summary>필수 데이터 설정 데이터가 0개라면 에러</summary>
    public bool RequiredColumn { get; set; }
    /// <summary>해당 데이터가 빈칸이라면 Default Value 설정</summary>
    public object DefaultValue { get; set; }
    /// <summary>커스텀 파서</summary>
    public Type CustomParser { get; set; }
    //Split 구분자
    /// <summary>
    /// Split 구분자
    /// </summary>
    public string Separator { get; set; }
    //머지를 하나로 합치는 여부
    /// <summary>
    /// 머지를 하나로 합치는 여부
    /// hp # 1  // hp # 2 라면
    /// 2           3
    /// 
    /// 2,3 로 들어간다.
    /// </summary>
    public bool MergedCells { get; set; }


    public ExcelParerAttribute(
        string columnName = null,
        bool ignore = false,
        bool requiredColumn = false,
        object defaultValue = null,
        Type customParser = null,
        string separator = ",",
        bool mergedCells = false)
    {
        this.ColumnName = columnName;
        this.Ignore = ignore;
        this.RequiredColumn = requiredColumn;
        this.DefaultValue = defaultValue;
        this.CustomParser = customParser;
        this.Separator = separator;
        this.MergedCells = mergedCells;
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

/// <summary>
/// 파싱 결과가 [Min, Max] 범위여야 함
/// </summary>
[AttributeUsage(AttributeTargets.Field)]
public class ValidateRangeAttribute : Attribute
{
    public double Min { get; }
    public double Max { get; }
    public ValidateRangeAttribute(double min, double max)
    {
        Min = min;
        Max = max;
    }
}

/// <summary>
/// 파싱 결과가 특정 정규식 패턴과 일치해야 함
/// </summary>
[AttributeUsage(AttributeTargets.Field)]
public class ValidateRegexAttribute : Attribute
{
    public string Pattern { get; }
    public ValidateRegexAttribute(string pattern)
    {
        Pattern = pattern;
    }
}