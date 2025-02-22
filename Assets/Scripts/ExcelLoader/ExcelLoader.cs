using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using UnityEngine;

public class ContainerFieldInfo
{
    public FieldInfo Field;
    public SheetBindingAttribute Binding;
}

public static class ExcelLoader
{
    public static void LoadAllExcelFiles<T>(T container, string folderPath) where T : class
    {
        if (!Directory.Exists(folderPath))
        {
            Debug.LogError($"[ExcelLoader] Folder not found: {folderPath}");
            return;
        }

        if (container == null)
        {
            container = (T)Activator.CreateInstance(typeof(T));
        }

        // 컨테이너의 필드와 SheetBindingAttribute 미리 캐싱
        //var containerFields = container.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance)
        //                                 .Select(f => new { Field = f, Binding = f.GetCustomAttribute<SheetBindingAttribute>() })
        //                                 .ToList();

        var containerFields = container.GetType()
    .GetFields(BindingFlags.Public | BindingFlags.Instance)
    .Select(f => new ContainerFieldInfo
    {
        Field = f,
        Binding = f.GetCustomAttribute<SheetBindingAttribute>()
    })
    .ToList();

        var files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
        foreach (var file in files)
        {
            string fileName = Path.GetFileName(file);
            if (fileName.StartsWith("~")) continue;
            LoadExcel(container, file, containerFields);
        }

        // 필수 바인딩 필드 체크
        foreach (var entry in containerFields)
        {
            if (entry.Binding != null && !entry.Binding.optional && entry.Field.GetValue(container) == null)
            {
                throw new Exception($"[ExcelLoader] Sheet not found for {entry.Field.Name}");
            }
        }
    }

    private static void LoadExcel<T>(T container, string filePath, List<ContainerFieldInfo> containerFields) where T : class
    {
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var ds = reader.AsDataSet();

            foreach (DataTable sheet in ds.Tables)
            {
                string rawSheet = sheet.TableName ?? "";
                if (rawSheet.StartsWith("~") || rawSheet.StartsWith("#")) continue;

                // 접두사 "!"가 있으면 제거하고 역전 처리를 위한 플래그 설정
                bool isColumnBased = false;
                if (rawSheet.StartsWith("!"))
                {
                    isColumnBased = true;
                    rawSheet = rawSheet.Substring(1);
                }
                string sheetName = rawSheet.Split('#')[0].Trim();

                // sheet와 매칭되는 필드를 미리 찾음
                var matchedFieldEntry = containerFields.FirstOrDefault(entry =>
                    entry.Field.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase) ||
                    (entry.Binding != null && entry.Binding.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase)));

                if (matchedFieldEntry != null)
                {
                    Debug.Log($"{sheetName}");
                    ParseSheetAndStore(container, sheet, matchedFieldEntry.Field, isColumnBased);
                }
            }
        }
    }

    /// <summary>
    /// isColumnBased에 따라 행/열 기반 파싱 메서드를 호출한 뒤,
    /// StoreInContainer로 넘겨주는 메서드
    /// </summary>
    private static void ParseSheetAndStore(object container, DataTable sheet, FieldInfo field, bool isColumnBased)
    {
        List<Dictionary<string, string>> dataList = ParseSheet(sheet, isColumnBased);

        // 이후 dataList를 컨테이너에 저장
        StoreInContainer(container, sheet, field, dataList);
    }

    private static List<Dictionary<string, string>> ParseSheet(DataTable sheet, bool isColumnBased)
    {
        List<Dictionary<string, string>> dataList = new();

        // primary: 헤더가 위치하는 축 (행 기반이면 첫 행, 열 기반이면 첫 열)
        // secondary: 데이터가 위치하는 축 (행 기반이면 행, 열 기반이면 열)
        int primaryCount = isColumnBased ? sheet.Rows.Count : sheet.Columns.Count;
        int secondaryCount = isColumnBased ? sheet.Columns.Count : sheet.Rows.Count;

        // 최소 2개(헤더 + 데이터)가 있어야 함.
        if (primaryCount <= 1)
        {
            Debug.LogWarning($"[ExcelLoader] Sheet {sheet.TableName} is empty or lacks enough {(isColumnBased ? "rows" : "columns")} for parsing.");
            return dataList;
        }

        // 첫 번째 행(또는 열)을 헤더로 사용
        var headerMap = new Dictionary<int, string>();
        for (int i = 0; i < primaryCount; i++)
        {
            // isColumnBased: 헤더는 첫 열(즉, sheet.Rows[i][0]), 그렇지 않으면 첫 행(즉, sheet.Rows[0][i])
            string head = isColumnBased
                ? sheet.Rows[i][0]?.ToString() ?? ""
                : sheet.Rows[0][i]?.ToString() ?? "";
            headerMap[i] = head;
        }

        // 헤더명을 그룹핑: baseName -> 인덱스 목록
        var grouped = new Dictionary<string, List<int>>();
        for (int i = 0; i < primaryCount; i++)
        {
            string rawHeader = headerMap[i];
            if (string.IsNullOrWhiteSpace(rawHeader)) continue;
            if (rawHeader.StartsWith("~") || rawHeader.StartsWith("#")) continue;

            string baseName = rawHeader.Split('#')[0].Trim();
            if (!grouped.ContainsKey(baseName))
                grouped[baseName] = new List<int>();
            grouped[baseName].Add(i);
        }

        // 데이터 처리: 데이터는 헤더 이후의 secondary 축(행 또는 열)부터 처리
        for (int j = 1; j < secondaryCount; j++)
        {
            var fieldValues = new Dictionary<string, string>();
            bool hasData = false;

            foreach (var kv in grouped)
            {
                string baseName = kv.Key;
                List<int> indices = kv.Value;
                List<string> parts = new();

                foreach (int i in indices)
                {
                    // isColumnBased: 데이터는 sheet.Rows[i][j] (헤더가 열에 있으므로, j번째 열의 데이터)
                    // 행 기반: 데이터는 sheet.Rows[j][i] (헤더가 행에 있으므로, j번째 행의 데이터)
                    string cellVal = isColumnBased
                        ? sheet.Rows[i][j]?.ToString() ?? ""
                        : sheet.Rows[j][i]?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(cellVal))
                    {
                        parts.Add(cellVal.Trim());
                        hasData = true;
                    }
                }
                fieldValues[baseName] = string.Join(",", parts);
            }
            if (!hasData)
                break;
            dataList.Add(fieldValues);
        }

        return dataList;
    }

    private static object ConvertAndValidate(string cellStr, FieldInfo field, DataTable sheet)
    {
        if (string.IsNullOrWhiteSpace(cellStr))
        {
            var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();
            return excelParer != null ? excelParer.DefaultValue : GetDefaultValue(field.FieldType);
        }

        object finalVal = null;
        try
        {
            var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();
            // 먼저 정적 ParseValue 메서드 시도
            var customParserValue = TryParseUsingStaticMethod(cellStr, field.FieldType);
            if (customParserValue != null)
            {
                finalVal = customParserValue;
            }
            else if (excelParer != null && excelParer.CustomParser != null)
            {
                ICustomParser parser = (ICustomParser)Activator.CreateInstance(excelParer.CustomParser);
                finalVal = parser.Parse(cellStr);
            }
            else if (typeof(ICustomParser).IsAssignableFrom(field.FieldType))
            {
                ICustomParser parser = (ICustomParser)Activator.CreateInstance(field.FieldType);
                finalVal = parser?.Parse(cellStr);
            }
            else if (field.FieldType.IsEnum)
            {
                if (Enum.TryParse(field.FieldType, cellStr, true, out object enResult))
                {
                    finalVal = enResult;
                }
                else
                {
                    Debug.LogError($"Enum Parse Error sheet : {sheet.Namespace} fieldName : {field.Name} {field.FieldType.Name} , cellString : {cellStr}");
                    return excelParer != null ? excelParer.DefaultValue : GetDefaultValue(field.FieldType);
                }
            }
            else if (field.FieldType.IsArray)
            {
                Type elemType = field.FieldType.GetElementType();
                var splitted = cellStr.Split(',').Select(s => Convert.ChangeType(s.Trim(), elemType)).ToArray();
                var arr = Array.CreateInstance(elemType, splitted.Length);
                splitted.CopyTo(arr, 0);
                finalVal = arr;
            }
            else if (field.FieldType.IsGenericType && field.FieldType.GetGenericTypeDefinition() == typeof(List<>))
            {
                var elemType = field.FieldType.GetGenericArguments()[0];
                var listObj = Activator.CreateInstance(field.FieldType) as System.Collections.IList;
                foreach (var part in cellStr.Split(','))
                {
                    string trimmed = part.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                        listObj.Add(Convert.ChangeType(trimmed, elemType));
                }
                finalVal = listObj;
            }
            else if (field.FieldType == typeof(Vector2))
            {
                finalVal = Vector2Parser.ParseValue(cellStr);
            }
            else if (field.FieldType == typeof(Vector3))
            {
                finalVal = Vector3Parser.ParseValue(cellStr);
            }
            else
            {
                finalVal = Convert.ChangeType(cellStr, field.FieldType);
            }
        }
        catch
        {
            Debug.LogError($"Convert Error Sheet {sheet.TableName} : {field.Name} {field.FieldType.Name} , cellString : {cellStr}");
            var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();
            return excelParer != null ? excelParer.DefaultValue : GetDefaultValue(field.FieldType);
        }

        // 범위 검증
        var rangeAttr = field.GetCustomAttribute<ValidateRangeAttribute>();
        if (rangeAttr != null)
        {
            double dVal = 0;
            try { dVal = Convert.ToDouble(finalVal); } catch { }
            if (dVal < rangeAttr.Min || dVal > rangeAttr.Max)
                throw new Exception($"[ExcelLoader] sheet : {sheet.TableName} {field.Name}={dVal} out of range [{rangeAttr.Min},{rangeAttr.Max}]");
        }

        // 정규식 검증
        var regexAttr = field.GetCustomAttribute<ValidateRegexAttribute>();
        if (regexAttr != null)
        {
            string sVal = finalVal?.ToString() ?? "";
            if (!Regex.IsMatch(sVal, regexAttr.Pattern))
                throw new Exception($"[ExcelLoader] sheet : {sheet.TableName} {field.Name}='{sVal}' doesn't match pattern '{regexAttr.Pattern}'");
        }

        return finalVal;
    }

    private static object TryParseUsingStaticMethod(string value, Type targetType)
    {
        var method = targetType.GetMethod("ParseValue", BindingFlags.Public | BindingFlags.Static);
        if (method != null)
            return method.Invoke(null, new object[] { value });
        return null;
    }

    private static object GetDefaultValue(Type t)
    {
        if (t == typeof(string)) return "";
        if (t == typeof(int) || t == typeof(long) || t == typeof(short) || t == typeof(byte) ||
            t == typeof(uint) || t == typeof(ulong) || t == typeof(ushort))
            return 0;
        if (t == typeof(float)) return 0f;
        if (t == typeof(double)) return 0.0;
        if (t == typeof(decimal)) return 0m;
        if (t == typeof(bool)) return false;
        if (t == typeof(Vector2)) return Vector2.zero;
        if (t == typeof(Vector3)) return Vector3.zero;
        if (t.IsValueType) return Activator.CreateInstance(t);
        return null;
    }

    private static void StoreInContainer(object container, DataTable sheet, FieldInfo parentField, List<Dictionary<string, string>> dataList)
    {
        Type dataType = GetDataType(parentField);
        var bindAttr = parentField.GetCustomAttribute<SheetBindingAttribute>();

        // dataType의 필드와 관련 어트리뷰트를 미리 캐싱 (이름 기준 매핑)
        var dataFields = dataType.GetFields(BindingFlags.Public | BindingFlags.Instance)
                                 .Select(f => new { Field = f, Parser = f.GetCustomAttribute<ExcelParerAttribute>(), MultiParser = f.GetCustomAttribute<MultiColumnParserAttribute>() })
                                 .ToDictionary(x => x.Field.Name, x => x);

        foreach (var data in dataList)
        {
            object instance = Activator.CreateInstance(dataType);
            object objectKey = null;

            // 단일 컬럼 데이터 매핑
            foreach (var kv in data)
            {
                if (dataFields.TryGetValue(kv.Key, out var fieldInfo))
                {
                    object fieldValue = ConvertAndValidate(kv.Value, fieldInfo.Field, sheet);
                    // 첫번째 값은 key로 활용할 수 있음.
                    if (objectKey == null) objectKey = fieldValue;
                    fieldInfo.Field.SetValue(instance, fieldValue);
                }
            }

            // 멀티 컬럼 파서 처리 (필드 단위)
            foreach (var kv in dataFields)
            {
                var mpAttr = kv.Value.MultiParser;
                if (mpAttr == null || mpAttr.ColumnNames == null || mpAttr.ColumnNames.Length == 0)
                    continue;

                bool isValid = mpAttr.ColumnNames.All(col => !string.IsNullOrWhiteSpace(col) && data.ContainsKey(col));
                if (!isValid) continue;

                var values = mpAttr.ColumnNames.Select(col => data[col]).ToArray();
                var mp = (IMultiColumnParser)Activator.CreateInstance(mpAttr.ParserType);
                kv.Value.Field.SetValue(instance, mp.Parse(values));
            }

            // key 결정 : Key() 메서드 우선, 없으면 첫번째 값
            object key = null;
            var keyMethod = dataType.GetMethod("Key");
            key = keyMethod != null ? keyMethod.Invoke(instance, null)?.ToString() : objectKey;

            FillBoundField(container, parentField, dataType, key, instance, bindAttr , sheet);
        }
    }

    private static Type GetDataType(FieldInfo field)
    {
        if (IsDictType(field.FieldType, out var keyType, out var valType))
            return valType;
        else if (field.FieldType.IsArray)
            return field.FieldType.GetElementType();
        else if (field.FieldType.IsGenericType && field.FieldType.GetGenericTypeDefinition() == typeof(List<>))
            return field.FieldType.GetGenericArguments()[0];
        return field.FieldType;
    }

    private static void FillBoundField(object container, FieldInfo field, Type dataType, object key, object dataList, SheetBindingAttribute bindAttr ,DataTable sheet)
    {
        if (IsDictType(field.FieldType, out var keyType, out var valType))
        {
            if (valType == dataType)
            {
                var dictVal = field.GetValue(container);
                if (dictVal == null)
                {
                    dictVal = Activator.CreateInstance(field.FieldType);
                    field.SetValue(container, dictVal);
                }
                var dictID = dictVal as System.Collections.IDictionary;
                if (dictID != null)
                {
                    if (dictID.Contains(key))
                    {
                        if (!bindAttr.skipDuplicates)
                            throw new Exception($"[ExcelLoader] sheet {sheet.TableName} Duplicate key {key} in dict field={field.Name}");
                    }
                    dictID[key] = dataList;
                }
            }
            else
            {
                Debug.LogWarning($"[ExcelLoader] sheet {sheet.TableName} field {field.Name}: dictionary ValueType != {dataType.Name}");
            }
        }
        else if (field.FieldType == dataType)
        {
            field.SetValue(container, dataList);
        }
        else if (IsListOfType(field.FieldType, dataType))
        {
            var listVal = field.GetValue(container) as System.Collections.IList;
            if (listVal == null)
            {
                listVal = Activator.CreateInstance(field.FieldType) as System.Collections.IList;
                field.SetValue(container, listVal);
            }
            listVal.Add(dataList);
        }
        else if (field.FieldType.IsArray && field.FieldType.GetElementType() == dataType)
        {
            var existingArray = field.GetValue(container) as Array;
            int existingLength = existingArray != null ? existingArray.Length : 0;
            Array newArray = Array.CreateInstance(dataType, existingLength + 1);
            if (existingArray != null)
                Array.Copy(existingArray, newArray, existingLength);
            newArray.SetValue(dataList, existingLength);
            field.SetValue(container, newArray);
        }
        else
        {
            Debug.LogWarning($"[ExcelLoader] sheet {sheet.TableName} field {field.Name} has [SheetBinding({dataType.Name})], but type mismatch? {field.FieldType}");
        }
    }

    private static bool IsListOfType(Type t, Type elem)
    {
        return t.IsGenericType && t.GetGenericTypeDefinition() == typeof(List<>) && t.GetGenericArguments()[0] == elem;
    }

    private static bool IsDictType(Type t, out Type keyType, out Type valType)
    {
        keyType = null;
        valType = null;
        if (!t.IsGenericType || t.GetGenericTypeDefinition() != typeof(Dictionary<,>))
            return false;
        var args = t.GetGenericArguments();
        keyType = args[0];
        valType = args[1];
        return true;
    }
}
