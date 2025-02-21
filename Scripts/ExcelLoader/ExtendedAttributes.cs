using System;

/// <summary>
/// (1) IgnoreParsing
/// 이 필드(또는 프로퍼티)는 엑셀 파싱에서 아예 무시
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class IgnoreParsingAttribute : Attribute
{
    // 필요한 옵션 있으면 추가
}

/// <summary>
/// (2) RequiredColumn
/// 이 필드에 매핑될 컬럼이 없으면 파싱 에러
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class RequiredColumnAttribute : Attribute
{
    // 옵션 필요하면 추가
}

/// <summary>
/// (3) ColumnName
/// 필드/프로퍼티 이름과 다른 엑셀 컬럼명을 직접 지정
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class ColumnNameAttribute : Attribute
{
    public string Name { get; }
    public ColumnNameAttribute(string name)
    {
        Name = name;
    }
}

/// <summary>
/// (4) ColumnIndex
/// 이 필드(또는 프로퍼티)는 엑셀의 N번째 컬럼과 매핑
/// (헤더 없이, 또는 고정 위치일 때 유용)
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class ColumnIndexAttribute : Attribute
{
    public int Index { get; }
    public ColumnIndexAttribute(int index)
    {
        Index = index;
    }
}

/// <summary>
/// (5) DefaultValue
/// 셀이 비어있거나 변환 실패 시 이 값을 사용
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class DefaultValueAttribute : Attribute
{
    public object Value { get; }
    public DefaultValueAttribute(object value)
    {
        Value = value;
    }
}

/// <summary>
/// (6-a) ValidateRange
/// 파싱 결과가 [Min, Max] 범위여야 함
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
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
/// (6-b) ValidateRegex
/// 파싱 결과가 특정 정규식 패턴과 일치해야 함
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class ValidateRegexAttribute : Attribute
{
    public string Pattern { get; }
    public ValidateRegexAttribute(string pattern)
    {
        Pattern = pattern;
    }
}
