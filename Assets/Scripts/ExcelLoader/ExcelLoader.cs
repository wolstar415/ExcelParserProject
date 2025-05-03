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
    static string[] excelExtensions = new[] { "*.xls", "*.xlsx", "*.xlsb", "*.csv" };

    public static T LoadExcelFile<T>(T container, string path) where T : class
    {
        if (!File.Exists(path))
        {
            Debug.LogError($"[ExcelLoader] File not found: {path}");
            return container;
        }

        container ??= (T)Activator.CreateInstance(typeof(T));
        var containerFields = GetContainerFields(container);

        LoadExcel(container, path, containerFields);

        ValidateContainerFields(container, containerFields);

        return container;
    }

    public static T LoadAllExcelFiles<T>(T container, string folderPath) where T : class
    {
        if (!Directory.Exists(folderPath))
        {
            Debug.LogError($"[ExcelLoader] Folder not found: {folderPath}");
            return container;
        }

        container ??= (T)Activator.CreateInstance(typeof(T));
        var containerFields = GetContainerFields(container);

        var excelFiles = excelExtensions
            .SelectMany(ext => Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly))
            .Where(file => !Path.GetFileName(file).StartsWith("~"))
            .ToArray();

        foreach (var file in excelFiles)
        {
            LoadExcel(container, file, containerFields);
        }

        ValidateContainerFields(container, containerFields);

        return container;
    }

    private static List<ContainerFieldInfo> GetContainerFields<T>(T container) where T : class
    {
        return container.GetType()
            .GetFields(BindingFlags.Public | BindingFlags.Instance)
            .Select(f => new ContainerFieldInfo
            {
                Field = f,
                Binding = f.GetCustomAttribute<SheetBindingAttribute>()
            })
            .ToList();
    }

    private static void ValidateContainerFields<T>(T container, List<ContainerFieldInfo> containerFields) where T : class
    {
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
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var ds = reader.AsDataSet();

        foreach (DataTable sheet in ds.Tables)
        {
            string rawSheet = sheet.TableName?.Trim() ?? "";
            if (string.IsNullOrEmpty(rawSheet) || rawSheet.StartsWith("~") || rawSheet.StartsWith("#")) continue;

            bool isColumnBased = rawSheet.StartsWith("!") || rawSheet.StartsWith("*");
            if (isColumnBased) rawSheet = rawSheet[1..];

            string sheetName = rawSheet.Split('#')[0].Trim();

            var matchedFieldEntrys = containerFields
                .Where(entry => entry.Binding?.SheetName?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true ||
                                entry.Field.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var entry in matchedFieldEntrys)
            {
                bool _columnBase = isColumnBased && entry.Binding?.isColumnBased == true;
                ParseSheetAndStore(container, sheet, entry.Field, isColumnBased || _columnBase);
            }
        }
    }
    private static void ParseSheetAndStore(object container, DataTable sheet, FieldInfo field, bool isColumnBased)
    {
        var dataList = ParseSheet(sheet, isColumnBased);
        StoreInContainer(container, sheet, field, dataList);
    }

    private static List<Dictionary<string, List<string>>> ParseSheet(DataTable sheet, bool isColumnBased)
    {
        var dataList = new List<Dictionary<string, List<string>>>();

        int primaryCount = isColumnBased ? sheet.Rows.Count : sheet.Columns.Count;
        int secondaryCount = isColumnBased ? sheet.Columns.Count : sheet.Rows.Count;

        if (primaryCount <= 1)
        {
            Debug.LogWarning($"[ExcelLoader] Sheet {sheet.TableName} is empty or lacks enough {(isColumnBased ? "rows" : "columns")} for parsing.");
            return dataList;
        }

        int startIndex = 0;
        for (int i = 0; i < secondaryCount; i++)
        {
            string head = isColumnBased ? sheet.Rows[0][i]?.ToString() ?? "" : sheet.Rows[i][0]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(head) && (head.StartsWith("//") || head.StartsWith("##")))
            {
                startIndex++;
                continue;
            }
            break;
        }

        if (primaryCount <= startIndex + 1)
        {
            Debug.LogWarning($"[ExcelLoader] Sheet {sheet.TableName} is empty or lacks enough {(isColumnBased ? "rows" : "columns")} for parsing.");
            return dataList;
        }

        var headerMap = new Dictionary<int, string>();
        for (int i = 0; i < primaryCount; i++)
        {
            string head = isColumnBased ? sheet.Rows[i][startIndex]?.ToString() ?? "" : sheet.Rows[startIndex][i]?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(head)) headerMap[i] = head;
        }

        var grouped = headerMap
            .Where(kv => !string.IsNullOrWhiteSpace(kv.Value) && !kv.Value.StartsWith("~") && !kv.Value.StartsWith("#"))
            .GroupBy(kv => kv.Value.Split('#')[0].Trim())
            .ToDictionary(g => g.Key, g => g.Select(kv => kv.Key).ToList());

        for (int j = startIndex + 1; j < secondaryCount; j++)
        {
            string head = isColumnBased ? sheet.Rows[0][j]?.ToString() ?? "" : sheet.Rows[j][0]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(head) && (head.StartsWith("//") || head.StartsWith("##"))) continue;

            var fieldValues = new Dictionary<string, List<string>>();
            bool hasData = false;

            foreach (var kv in grouped)
            {
                var values = kv.Value.Select(i => isColumnBased ? sheet.Rows[i][j]?.ToString() ?? "" : sheet.Rows[j][i]?.ToString() ?? "").ToList();
                if (values.Any(v => !string.IsNullOrWhiteSpace(v))) hasData = true;
                fieldValues[kv.Key] = values;
            }

            if (!hasData) break;
            dataList.Add(fieldValues);
        }

        return dataList;
    }

    private static object ConvertAndValidate(List<string> cellStrList, FieldInfo field, DataTable sheet)
    {
        ExcelParerAttribute excelParer = field.GetCustomAttribute<ExcelParerAttribute>();

        string cellStr = "";

        string separator = ",";

        if (excelParer != null)
        {
            separator = excelParer.Separator;
        }

        if (cellStrList != null && cellStrList.Count > 0)
        {
            cellStr = cellStrList[0];
        }

        if (string.IsNullOrWhiteSpace(cellStr))
        {
            return excelParer != null ? excelParer.DefaultValue : GetDefaultValue(field.FieldType);
        }

        if (excelParer != null && excelParer.MergedCells)
        {
            string mergedCells = string.Join(separator, cellStrList.Where(item => !string.IsNullOrEmpty(item)));

            if (string.IsNullOrWhiteSpace(mergedCells) == false)
            {
                cellStr += $"{separator}{mergedCells}";
            }

            cellStrList = new List<string> { cellStr };

        }

        object finalVal = null;
        try
        {
            if (field.FieldType.IsArray)
            {
                Type elemType = field.FieldType.GetElementType();

                if (cellStrList != null && cellStrList.Count > 1)
                {
                    var splitted = cellStrList.Select(s => TryParseCellStr(s.Trim(), elemType, excelParer)).ToArray();
                    var arr = Array.CreateInstance(elemType, splitted.Length);
                    splitted.CopyTo(arr, 0);
                    finalVal = arr;
                }
                else
                {
                    var splitted = cellStr.Split(separator).Select(s => TryParseCellStr(s.Trim(), elemType, excelParer)).ToArray();
                    var arr = Array.CreateInstance(elemType, splitted.Length);
                    splitted.CopyTo(arr, 0);
                    finalVal = arr;
                }


            }
            else if (field.FieldType.IsGenericType && field.FieldType.GetGenericTypeDefinition() == typeof(List<>))
            {
                var elemType = field.FieldType.GetGenericArguments()[0];

                var listObj = Activator.CreateInstance(field.FieldType) as System.Collections.IList;

                if (cellStrList != null && cellStrList.Count > 1)
                {
                    foreach (var part in cellStrList)
                    {
                        string trimmed = part.Trim();
                        listObj.Add(TryParseCellStr(trimmed, elemType, excelParer));
                    }
                }
                else
                {
                    foreach (var part in cellStr.Split(separator))
                    {
                        string trimmed = part.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                            listObj.Add(TryParseCellStr(trimmed, elemType, excelParer));
                    }
                }


                finalVal = listObj;
            }
            else
            {
                finalVal = TryParseCellStr(cellStrList, field.FieldType, excelParer);
            }
        }
        catch
        {
            Debug.LogError($"Convert Error Sheet {sheet.TableName} : {field.Name} {field.FieldType.Name} , cellString : {cellStr}");
            //var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();
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

    private static object TryParseCellStr(string cellStr, Type type, ExcelParerAttribute excelParer = null)
    {
        var customParserValue = TryParseUsingStaticMethod(cellStr, type);

        if (customParserValue != null)
        {
            return customParserValue;
        }
        else if (excelParer != null && excelParer.CustomParser != null)
        {
            ICustomParser parser = (ICustomParser)Activator.CreateInstance(excelParer.CustomParser);
            return parser.Parse(cellStr);
        }
        else if (typeof(ICustomParser).IsAssignableFrom(type))
        {
            ICustomParser parser = (ICustomParser)Activator.CreateInstance(type);
            return parser?.Parse(cellStr);
        }
        else if (type.IsEnum)
        {
            if (Enum.TryParse(type, cellStr, true, out object enResult))
            {
                return enResult;
            }
            else
            {
                return excelParer != null ? excelParer.DefaultValue : GetDefaultValue(type);
            }
        }
        else if (type == typeof(Vector2))
        {
            return Vector2Parser.ParseValue(cellStr);
        }
        else if (type == typeof(Vector3))
        {
            return Vector3Parser.ParseValue(cellStr);
        }
        else
        {
            return Convert.ChangeType(cellStr, type);
        }
    }

    private static object TryParseCellStr(List<string> cellStrList, Type type, ExcelParerAttribute excelParer = null)
    {
        string separator = ",";

        string cellStr = "";

        if (excelParer != null)
        {
            separator = excelParer.Separator;
        }

        if (cellStrList != null && cellStrList.Count > 0)
        {
            cellStr = cellStrList[0];
        }

        if (excelParer != null && excelParer.MergedCells)
        {
            return string.Join(separator, cellStrList.Where(item => !string.IsNullOrEmpty(item)));
        }

        return TryParseCellStr(cellStr, type, excelParer);
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

    private static void StoreInContainer(object container, DataTable sheet, FieldInfo parentField, List<Dictionary<string, List<string>>> dataList)
    {
        Type dataType = GetDataType(parentField);

        if (dataType.IsGenericType && dataType.GetGenericTypeDefinition() == typeof(List<>))
        {
            dataType = dataType.GetGenericArguments()[0];

        }
        else if (dataType.IsArray)
        {
            dataType = dataType.GetElementType();
        }

        var bindAttr = parentField.GetCustomAttribute<SheetBindingAttribute>();

        var dataFields = dataType.GetFields(BindingFlags.Public | BindingFlags.Instance)
            .Select(f => new
            {
                Field = f,
                Parser = f.GetCustomAttribute<ExcelParerAttribute>(),
                MultiParser = f.GetCustomAttribute<MultiColumnParserAttribute>()
            })
            .Where(x => x.Parser == null || !x.Parser.Ignore) // Ignore가 true인 경우 제외
            .ToList();
        foreach (var data in dataList)
        {
            object instance = Activator.CreateInstance(dataType);
            object objectKey = null;

            foreach (var kv in data)
            {
                if (dataFields != null && dataFields.Count > 0)
                {
                    foreach (var fieldInfo in dataFields)
                    {
                        string name = fieldInfo.Field.Name;

                        if (fieldInfo.Parser != null && string.IsNullOrEmpty(fieldInfo.Parser.ColumnName) == false)
                            name = fieldInfo.Parser.ColumnName;

                        if (string.Equals(name, kv.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            object fieldValue = ConvertAndValidate(kv.Value, fieldInfo.Field, sheet);

                            if (objectKey == null) objectKey = fieldValue;
                            fieldInfo.Field.SetValue(instance, fieldValue);
                            break;
                        }
                    }
                }
            }

            foreach (var kv in dataFields)
            {
                var mpAttr = kv.MultiParser;
                if (mpAttr == null || mpAttr.ColumnNames == null || mpAttr.ColumnNames.Length == 0)
                    continue;

                bool isValid = mpAttr.ColumnNames.All(col => !string.IsNullOrWhiteSpace(col) && data.ContainsKey(col));
                if (!isValid) continue;

                var values = mpAttr.ColumnNames.Select(col => ((data[col] == null || data[col].Count == 0) ? "" : data[col][0])).ToArray();
                var mp = (IMultiColumnParser)Activator.CreateInstance(mpAttr.ParserType);
                kv.Field.SetValue(instance, mp.Parse(values));
            }

            object key = null;
            var keyMethod = dataType.GetMethod("Key");
            key = keyMethod != null ? (keyMethod.Invoke(instance, null)) : objectKey;

            FillBoundField(container, parentField, dataType, key, instance, bindAttr, sheet);
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

    private static void FillBoundField(object container, FieldInfo field, Type dataType, object key, object dataList, SheetBindingAttribute bindAttr, DataTable sheet)
    {
        if (IsDictType(field.FieldType, out var keyType, out var valType))
        {

            var dictVal = field.GetValue(container);
            if (dictVal == null)
            {
                dictVal = Activator.CreateInstance(field.FieldType);
                field.SetValue(container, dictVal);
            }
            var dictID = dictVal as System.Collections.IDictionary;
            if (dictID == null)
                return;

            if (dictID.Contains(key))
            {
                object existingValue = dictID[key];
                if (existingValue != null && (existingValue is System.Collections.IList || existingValue.GetType().IsArray))
                {
                    if (existingValue is System.Collections.IList list)
                    {
                        list.Add(dataList);
                    }
                    else if (existingValue.GetType().IsArray)
                    {
                        Array array = (Array)existingValue;
                        int len = array.Length;
                        Array newArray = Array.CreateInstance(valType.GetElementType(), len + 1);
                        Array.Copy(array, newArray, len);
                        newArray.SetValue(dataList, len);
                        dictID[key] = newArray;
                    }
                }
                else
                {
                    if (bindAttr != null && !bindAttr.skipDuplicates)
                    {
                        throw new Exception($"[ExcelLoader] Duplicate key {key} in dict field={field.Name}");
                    }
                    else
                    {
                        dictID[key] = dataList;
                    }
                }
            }
            else
            {
                if (valType.IsArray)
                {
                    Array newArray = Array.CreateInstance(valType.GetElementType(), 1);
                    newArray.SetValue(dataList, 0);
                    dictID[key] = newArray;
                }
                else if (valType.IsGenericType && valType.GetGenericTypeDefinition() == typeof(List<>))
                {
                    var list = Activator.CreateInstance(valType) as System.Collections.IList;
                    list.Add(dataList);
                    dictID[key] = list;
                }
                else
                {
                    dictID[key] = dataList;
                }
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
