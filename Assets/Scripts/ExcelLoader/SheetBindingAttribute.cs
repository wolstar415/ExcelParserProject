using System;

/// <summary>
/// 필드에 붙여서, "이 필드는 어떤 시트와 매핑되는지"를 지정.
/// 예) [SheetBinding("UnitData", skipIfSheetNotFound=true, optional=true, skipDuplicates=true)]
/// </summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class SheetBindingAttribute : Attribute
{
    /// <summary>연결할 시트(또는 dataType.Name)</summary>
    public string SheetName { get; }

    /// <summary>데이터가 0개(또는 시트가 비었을 때) 허용할지(false면 경고/에러)</summary>
    public bool optional { get; set; }

    /// <summary>딕셔너리 중복 키가 발생하면 스킵할지(false면 예외)</summary>
    public bool skipDuplicates { get; set; }

    /// <summary> Column 으로 저장하는 방식</summary>
    public bool isColumnBased { get; set; }


    public SheetBindingAttribute(
        string sheetName = null,
        bool optional = true,
        bool skipDuplicates = false,
        bool isColumnBased = false)
    {
        this.SheetName = sheetName;
        this.optional = optional;
        this.skipDuplicates = skipDuplicates;
        this.isColumnBased = isColumnBased;
    }
}
