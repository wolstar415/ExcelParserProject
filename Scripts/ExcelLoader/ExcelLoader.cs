using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using UnityEngine;

public static class ExcelLoader
{
    public static void LoadAllExcelFiles(object container, string folderPath)
    {
        if (!Directory.Exists(folderPath))
        {
            Debug.LogError($"[ExcelLoader] Folder not found: {folderPath}");
            return;
        }

        var files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
        foreach (var file in files)
        {
            string fileName = Path.GetFileName(file);
            if (fileName.StartsWith("~"))
            {
                continue;
            }
            LoadExcel(container, file);
        }

        var cFields = container.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);

        foreach (var field in cFields)
        {
            var bindAttr = field.GetCustomAttribute<SheetBindingAttribute>();

            if (bindAttr != null && bindAttr.optional == false && field.GetValue(container) == null)
            {
                throw new Exception($"[ExcelLoader] Sheet not found for {field.Name}");
            }
        }
    }

    private static void LoadExcel(object container, string filePath)
    {
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var ds = reader.AsDataSet();

            int count = 0;
            foreach (DataTable sheet in ds.Tables)
            {
                string rawSheet = sheet.TableName ?? "";
                if (rawSheet.StartsWith("~") || rawSheet.StartsWith("#"))
                {
                    continue;
                }
                string sheetName = rawSheet.Split('#')[0].Trim();

                var cFields = container.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);

                foreach (var field in cFields)
                {
                    var bindAttr = field.GetCustomAttribute<SheetBindingAttribute>();

                    if (field.Name == sheetName || (bindAttr != null && bindAttr.SheetName == sheetName))
                    {
                        ParseSheetAndStore(container, sheet, field);
                    }
                }
                count++;
            }
        }
    }

    private static void ParseSheetAndStore(object container, DataTable sheet, FieldInfo field)
    {
        int rowCount = sheet.Rows.Count;

        if (rowCount <= 1)
        {
            Debug.LogWarning($"[ExcelLoader] Sheet {sheet.TableName} is empty.");
        }

        int colCount = sheet.Columns.Count;

        Dictionary<int, string> headerMap = new Dictionary<int, string>();
        for (int c = 0; c < colCount; c++)
        {
            string head = sheet.Rows[0][c]?.ToString() ?? "";
            headerMap[c] = head;
        }

        Dictionary<string, List<int>> groupedCols = new Dictionary<string, List<int>>();
        for (int c = 0; c < colCount; c++)
        {
            string rawHeader = headerMap[c];
            if (string.IsNullOrWhiteSpace(rawHeader)) continue;
            if (rawHeader.StartsWith("~") || rawHeader.StartsWith("#")) continue;

            string baseName = rawHeader.Split('#')[0].Trim();  // "bonusDamage#1" => "bonusDamage"
            if (!groupedCols.ContainsKey(baseName))
                groupedCols[baseName] = new List<int>();
            groupedCols[baseName].Add(c);
        }

        List<Dictionary<string, string>> dataList = new();

        for (int r = 1; r < rowCount; r++)
        {
            Dictionary<string, string> fieldValues = new Dictionary<string, string>();

            bool _isBreak = true;
            foreach (var kv in groupedCols)
            {
                string baseName = kv.Key;
                List<int> cols = kv.Value;  // 여러 col index

                List<string> parts = new List<string>();

                foreach (int cIdx in cols)
                {
                    string cellVal = sheet.Rows[r][cIdx]?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(cellVal))
                        parts.Add(cellVal.Trim());
                }
                if (parts.Count > 0)
                    _isBreak = false;

                fieldValues.Add(baseName, string.Join(",", parts));
            }

            if (_isBreak)
                break;

            dataList.Add(fieldValues);
        }

        StoreInContainer(container, sheet, field, dataList);
    }

    private static object ConvertAndValidate(string cellStr, FieldInfo field)
    {
        if (string.IsNullOrWhiteSpace(cellStr))
        {
            var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();

            if (excelParer != null) return excelParer.DefaultValue;
            return GetDefaultValue(field.FieldType);
        }

        object finalVal = null;
        try
        {
            var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();

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

                if (parser != null)
                    finalVal = parser.Parse(cellStr);
            }
            else if (field.FieldType.IsEnum)
            {
                if (Enum.TryParse(field.FieldType, cellStr, true, out object enResult))
                {
                    finalVal = enResult;
                }
                else
                {
                    Debug.LogError($" Enum Parse Error fieldName : {field.Name} {field.FieldType.Name} , cellString : {cellStr}");

                    if (excelParer != null) return excelParer.DefaultValue;

                    return GetDefaultValue(field.FieldType);
                }
            }
            else if (field.FieldType.IsArray)
            {
                Type elemType = field.FieldType.GetElementType();
                var splitted = cellStr.Split(',')
                    .Select(s => Convert.ChangeType(s.Trim(), elemType))
                    .ToArray();
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
                    {
                        listObj.Add(Convert.ChangeType(trimmed, elemType));
                    }
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

            if (finalVal == null)
            {
                finalVal = Convert.ChangeType(cellStr, field.FieldType);
            }
        }
        catch
        {
            Debug.LogError($" Convert Error : {field.Name} {field.FieldType.Name} , cellString : {cellStr}");

            var excelParer = field.GetCustomAttribute<ExcelParerAttribute>();

            if (excelParer != null) return excelParer.DefaultValue;

            return GetDefaultValue(field.FieldType);
        }

        var rangeAttr = field.GetCustomAttribute<ValidateRangeAttribute>();
        if (rangeAttr != null)
        {
            double dVal = 0;
            try { dVal = Convert.ToDouble(finalVal); } catch { }
            if (dVal < rangeAttr.Min || dVal > rangeAttr.Max)
            {
                throw new Exception($"[ExcelLoader] {field.Name}={dVal} out of range [{rangeAttr.Min},{rangeAttr.Max}]");
            }
        }

        var regexAttr = field.GetCustomAttribute<ValidateRegexAttribute>();
        if (regexAttr != null)
        {
            string sVal = finalVal?.ToString() ?? "";
            if (!Regex.IsMatch(sVal, regexAttr.Pattern))
            {
                throw new Exception($"[ExcelLoader] {field.Name}='{sVal}' doesn't match pattern '{regexAttr.Pattern}'");
            }
        }

        return finalVal;
    }

    private static object TryParseUsingStaticMethod(string value, Type targetType)
    {
        var method = targetType.GetMethod("ParseValue", BindingFlags.Public | BindingFlags.Static);
        if (method != null)
        {
            return method.Invoke(null, new object[] { value });
        }
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

    private static void StoreInContainer(object container,
                                     DataTable sheet,
                                     FieldInfo parentField,
                                     List<Dictionary<string, string>> dataList)
    {
        var fieldType = GetType(parentField);

        var bindAttr = parentField.GetCustomAttribute<SheetBindingAttribute>();

        var fields = fieldType.GetFields(BindingFlags.Public | BindingFlags.Instance);

        foreach (var data in dataList)
        {
            object instance = Activator.CreateInstance(fieldType);
            object objectKey = null;


            foreach (var field in fields)
            {
                var excelParser = parentField.GetCustomAttribute<ExcelParerAttribute>();

                if (excelParser != null && excelParser.Ignore)
                    continue;

                string name = field.Name;

                if (excelParser != null && excelParser.ColumnName != null)
                {
                    name = excelParser.ColumnName;
                }

                if (excelParser != null && excelParser.RequiredColumn && data.ContainsKey(name) == false)
                {
                    throw new Exception($"[ExcelLoader] Required column '{name}' not found for {sheet.TableName}");
                }

                foreach (var keyValue in data)
                {
                    string baseName = keyValue.Key;
                    string value = keyValue.Value;

                    if (name == baseName)
                    {
                        var fieldValue = ConvertAndValidate(value, field);

                        if (objectKey == null)
                        {
                            objectKey = fieldValue;
                        }

                        field.SetValue(instance, fieldValue);
                    }
                }
            }


            foreach (var field in fields)
            {
                var multiParser = field.GetCustomAttribute<MultiColumnParserAttribute>();

                if (multiParser == null || multiParser.ColumnNames == null || multiParser.ColumnNames.Length == 0)
                    continue;

                bool isbreak = false;
                foreach (var item in multiParser.ColumnNames)
                {
                    if (string.IsNullOrWhiteSpace(item))
                    {
                        isbreak = true;
                        break;
                    }
                    if (data.ContainsKey(item) == false)
                    {
                        isbreak = true;
                        break;
                    }
                }

                if (isbreak)
                    continue;

                List<string> values = new List<string>(multiParser.ColumnNames.Length);

                foreach (var item in multiParser.ColumnNames)
                {
                    values.Add(data[item]);
                }

                var mp = (IMultiColumnParser)Activator.CreateInstance(multiParser.ParserType);
                field.SetValue(instance, mp.Parse(values.ToArray()));
            }

            object key = null;

            var keyMethod = fieldType.GetMethod("Key");
            if (keyMethod != null)
            {
                key = keyMethod.Invoke(instance, null)?.ToString();
            }
            else
            {
                key = objectKey;
            }

            FillBoundField(container, parentField, fieldType, key, instance, bindAttr);
        }

    }

    private static Type GetType(FieldInfo field)
    {
        if (IsDictType(field.FieldType, out var keyType, out var valType))
        {
            return valType;
        }
        else if (field.FieldType.IsArray)
        {
            return field.FieldType.GetElementType();
        }

        else if (field.FieldType.IsGenericType && field.FieldType.GetGenericTypeDefinition() == typeof(List<>))
        {
            return field.FieldType.GetGenericArguments()[0];
        }

        return field.FieldType;
    }

    private static void FillBoundField(object container,
                                   FieldInfo field,
                                   Type dataType,
                                   object key,
                                   object dataList,
                                   SheetBindingAttribute bindAttr)
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
                        if (bindAttr != null && bindAttr.skipDuplicates)
                        {
                        }
                        else
                        {
                            throw new Exception($"[ExcelLoader] Duplicate key {key} in dict field={field.Name}");
                        }
                    }
                    dictID[key] = dataList;
                }
            }
            else
            {
                Debug.LogWarning($"[ExcelLoader] field {field.Name}: dictionary ValueType != {dataType.Name}");
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
                var newList = Activator.CreateInstance(field.FieldType);
                field.SetValue(container, newList);
                listVal = newList as System.Collections.IList;
            }
            listVal.Add(dataList);

        }
        else if (field.FieldType.IsArray && field.FieldType.GetElementType() == dataType)
        {
            var existingArray = field.GetValue(container) as Array;
            int existingLength = existingArray != null ? existingArray.Length : 0;
            int newLength = existingLength + 1;
            Array newArray = Array.CreateInstance(dataType, newLength);

            if (existingArray != null)
            {
                Array.Copy(existingArray, newArray, existingLength);
            }

            newArray.SetValue(dataList, existingLength);

            field.SetValue(container, newArray);
        }
        else
        {
            string name = field.FieldType.Name;

            Debug.LogWarning($"[ExcelLoader] field {field.Name} has [SheetBinding({name})], but type mismatch? {field.FieldType}");
        }
    }

    private static bool IsListOfType(Type t, Type elem)
    {
        if (!t.IsGenericType) return false;
        if (t.GetGenericTypeDefinition() != typeof(List<>)) return false;
        return t.GetGenericArguments()[0] == elem;
    }

    private static bool IsDictType(Type t, out Type keyType, out Type valType)
    {
        keyType = null;
        valType = null;

        if (!t.IsGenericType) return false;
        if (t.GetGenericTypeDefinition() != typeof(Dictionary<,>)) return false;

        var args = t.GetGenericArguments(); // [K, V]
        keyType = args[0];
        valType = args[1];
        return true;
    }
}