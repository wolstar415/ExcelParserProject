using UnityEngine;
using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelDataReader;

/// <summary>
/// 종합 로더
/// 1) SheetBindingAttribute로 시트-필드 연결
/// 2) 없으면 fallback(필드 타입으로 저장)
/// 3) multi-column (#1, #2) 병합
/// 4) ExtendedAttributes(1~6) 처리
/// 5) Dict 중복 key 시 skip or exception
/// 6) Key() 없는 경우 fallback
/// </summary>
public static class ExcelLoader
{
    public static void LoadAllExcelFiles(object container, string folderPath)
    {
        if (!Directory.Exists(folderPath))
        {
            Debug.LogError($"[ExcelLoader] Folder not found: {folderPath}");
            return;
        }

        // XLS 인코딩 (필요 없으면 주석)
        //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
        foreach (var file in files)
        {
            string fileName = Path.GetFileName(file);
            if (fileName.StartsWith("~"))
            {
                Debug.Log($"[ExcelLoader] Skip file ~: {fileName}");
                continue;
            }
            LoadExcel(container, file);
        }
    }

    private static void LoadExcel(object container, string filePath)
    {
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var ds = reader.AsDataSet();
            foreach (DataTable sheet in ds.Tables)
            {
                string rawSheet = sheet.TableName ?? "";
                if (rawSheet.StartsWith("~"))
                {
                    Debug.Log($"[ExcelLoader] Skip sheet ~: {rawSheet}");
                    continue;
                }
                // '#' 이후 무시
                string sheetName = rawSheet.Split('#')[0].Trim();

                Type dataType = Type.GetType(sheetName);
                if (dataType == null)
                {
                    // SheetBinding에서 skipIfSheetNotFound?
                    // -> 이건 container 필드 쪽에서 관리해야 함
                    continue;
                }
                ParseSheet(container, sheet, dataType);
            }
        }
    }

    /// <summary>
    /// 시트 -> dataType 파싱
    /// 1) multi-col (bonusDamage#1, #2, #3)
    /// 2) ExtendedAttributes
    /// 3) Key() or fallback
    /// -> dataList & dataDict
    /// -> StoreInContainer
    /// </summary>
    private static void ParseSheet(object container, DataTable sheet, Type dataType)
    {
        int rowCount = sheet.Rows.Count;
        if (rowCount <= 1)
        {
            Debug.LogWarning($"[ExcelLoader] Sheet {sheet.TableName} is empty.");
        }

        int colCount = sheet.Columns.Count;

        // (A) 헤더: index -> rawHeader
        // We'll do "bonusDamage#1,#2" grouping
        Dictionary<int, string> headerMap = new Dictionary<int, string>();
        for (int c = 0; c < colCount; c++)
        {
            string head = sheet.Rows[0][c]?.ToString() ?? "";
            headerMap[c] = head;
        }

        // groupedCols : baseName -> List<int> (column indexes)
        // ex) "bonusDamage#1" -> baseName="bonusDamage"
        //     "bonusDamage#2" -> same baseName
        Dictionary<string, List<int>> groupedCols = new Dictionary<string, List<int>>();
        for (int c = 0; c < colCount; c++)
        {
            string raw = headerMap[c];
            if (string.IsNullOrWhiteSpace(raw)) continue;
            if (raw.StartsWith("~") || raw.StartsWith("#")) continue;

            string baseName = raw.Split('#')[0].Trim();  // "bonusDamage#1" => "bonusDamage"
            if (!groupedCols.ContainsKey(baseName))
                groupedCols[baseName] = new List<int>();
            groupedCols[baseName].Add(c);
        }

        // (B) 파싱 대상 필드
        var fields = dataType.GetFields(BindingFlags.Public | BindingFlags.Instance);

        List<object> dataList = new List<object>();
        Dictionary<string, object> dataDict = new Dictionary<string, object>(); // Key() or fallback

        for (int r = 1; r < rowCount; r++)
        {
            object instance = Activator.CreateInstance(dataType);

            // 각 baseName -> merge => field
            foreach (var kv in groupedCols)
            {
                string baseName = kv.Key;
                var indexes = kv.Value;  // 여러 col index

                // IgnoreParsing?
                var field = fields.FirstOrDefault(f => f.Name == baseName);
                if (field == null)
                    continue; // no matching field name

                if (field.GetCustomAttribute<IgnoreParsingAttribute>() != null)
                    continue; // skip

                // 병합
                List<string> parts = new List<string>();
                foreach (int cIdx in indexes)
                {
                    string cellVal = (r < rowCount) ? sheet.Rows[r][cIdx]?.ToString() : "";
                    if (!string.IsNullOrWhiteSpace(cellVal))
                        parts.Add(cellVal.Trim());
                }
                // "1,2,3"
                string merged = string.Join(",", parts);

                // 변환 + validation
                object finalVal = ConvertAndValidate(merged, field);

                // setValue
                field.SetValue(instance, finalVal);
            }

            // Key() or fallback
            string key = null;
            var keyMethod = dataType.GetMethod("Key");
            if (keyMethod != null)
            {
                key = keyMethod.Invoke(instance, null)?.ToString();
            }
            else
            {
                // fallback: groupedCols 중 가장 작은 colIndex
                if (groupedCols.Count > 0)
                {
                    var firstPair = groupedCols.OrderBy(x => x.Value.Min()).First();
                    var fallbackField = fields.FirstOrDefault(f => f.Name == firstPair.Key);
                    if (fallbackField != null)
                    {
                        object val = fallbackField.GetValue(instance);
                        if (val != null) key = val.ToString();
                    }
                }
            }
            if (!string.IsNullOrEmpty(key))
            {
                if (dataDict.ContainsKey(key))
                {
                    // default: throw
                    // but skip if needed => handled later by skipDuplicates
                    dataDict[key] = instance; // or skip
                }
                else
                {
                    dataDict[key] = instance;
                }
            }
            dataList.Add(instance);
        }

        // (C) StoreInContainer
        StoreInContainer(container, dataType, dataList, dataDict);
        Debug.Log($"[ExcelLoader] Loaded {dataList.Count} rows for {dataType.Name} from sheet {sheet.TableName}");
    }

    /// <summary>
    /// ExtendedAttributes 처리( DefaultValue, RequiredColumn, ValidateRange, ValidateRegex ) 등
    /// ColumnIndex / ColumnName / RequiredColumn 은 헤더맵에서 결정할 때 사용
    /// 여기서는 "merged" 문자열 -> 최종 값 변환 + ValidateRange/Regex
    /// </summary>
    private static object ConvertAndValidate(string cellStr, FieldInfo field)
    {
        // DefaultValue check
        if (string.IsNullOrWhiteSpace(cellStr))
        {
            // default
            var defAttr = field.GetCustomAttribute<DefaultValueAttribute>();
            if (defAttr != null) return defAttr.Value;
            return GetDefaultValue(field.FieldType);
        }

        // Try parse
        object finalVal;
        try
        {
            // Enum?
            if (field.FieldType.IsEnum)
            {
                if (Enum.TryParse(field.FieldType, cellStr, true, out object enResult))
                {
                    finalVal = enResult;
                }
                else
                {
                    Debug.LogError($" Enum Parse Error fieldName : {field.FieldType.Name} , cellString : {cellStr}");
                    var defAttr = field.GetCustomAttribute<DefaultValueAttribute>();
                    if (defAttr != null) return defAttr.Value;
                    return GetDefaultValue(field.FieldType);
                }
            }
            else if (field.FieldType.IsArray)
            {
                // ex: "1,2,3" => int[]
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
                // e.g. List<int>
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
            else
            {
                // basic Convert
                finalVal = Convert.ChangeType(cellStr, field.FieldType);
            }
        }
        catch
        {
            var defAttr = field.GetCustomAttribute<DefaultValueAttribute>();
            if (defAttr != null) return defAttr.Value;
            return GetDefaultValue(field.FieldType);
        }

        // ValidateRange
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

        // ValidateRegex
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

    /// <summary>
    /// 타입별 기본값(0, "", false, null)
    /// </summary>
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
        if (t.IsValueType) return Activator.CreateInstance(t);
        return null;
    }

    private static void StoreInContainer(object container,
                                     Type dataType,
                                     List<object> dataList,
                                     Dictionary<string, object> dataDict)
    {
        // container의 public 필드 스캔
        var cFields = container.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);

        bool anyBoundField = false;

        // (1) SheetBinding 필드 먼저 처리
        foreach (var field in cFields)
        {
            var bindAttr = field.GetCustomAttribute<SheetBindingAttribute>();
            if (bindAttr != null && bindAttr.SheetName == dataType.Name)
            {
                anyBoundField = true;

                // optional=false + dataList.Count==0 => 경고/에러
                if (dataList.Count == 0 && !bindAttr.optional)
                {
                    Debug.LogWarning($"[ExcelLoader] sheet={dataType.Name} has no data, but 'optional=false' for field '{field.Name}'.");
                }

                // 실제 할당
                FillBoundField(container, field, dataType, dataList, dataDict, bindAttr);
            }
        }

        // (2) 만약 SheetBinding이 아예 없었다면 fallback
        if (!anyBoundField)
        {
            foreach (var field in cFields)
            {
                // IgnoreParsing?
                if (field.GetCustomAttribute<IgnoreParsingAttribute>() != null)
                    continue;

                // 단일 (fieldType == dataType)
                if (field.FieldType == dataType)
                {
                    if (dataList.Count > 0)
                    {
                        field.SetValue(container, dataList[0]);
                    }
                }
                // 리스트
                else if (IsListOfType(field.FieldType, dataType))
                {
                    var listObj = field.GetValue(container) as System.Collections.IList;
                    if (listObj != null)
                    {
                        foreach (var obj in dataList)
                            listObj.Add(obj);
                    }
                }
                // 딕셔너리 - key type 변환
                else if (IsDictType(field.FieldType, out var keyT, out var valT)
                         && valT == dataType)
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
                        foreach (var kvp in dataDict)
                        {
                            string strKey = kvp.Key;  // excel string key
                            object newKey = ConvertKeyString(strKey, keyT);
                            if (newKey == null)
                            {
                                Debug.LogWarning($"[ExcelLoader] skip key='{strKey}' can't parse as {keyT.Name}");
                                continue;
                            }
                            if (dictID.Contains(newKey))
                            {
                                // fallback => throw or skip
                                throw new Exception($"[ExcelLoader] Duplicate key {newKey} for field={field.Name}");
                            }
                            dictID[newKey] = kvp.Value;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// SheetBinding이 있는 필드 => 이 로직 우선
    /// </summary>
    private static void FillBoundField(object container,
                                       FieldInfo field,
                                       Type dataType,
                                       List<object> dataList,
                                       Dictionary<string, object> dataDict,
                                       SheetBindingAttribute bindAttr)
    {
        // (a) Dictionary<K,V>?
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
                    foreach (var kvp in dataDict)
                    {
                        // excel key => string
                        object convertedKey = ConvertKeyString(kvp.Key, keyType);
                        if (convertedKey == null)
                        {
                            Debug.LogWarning($"[ExcelLoader] skip invalid key='{kvp.Key}' for field={field.Name}");
                            continue;
                        }
                        if (dictID.Contains(convertedKey))
                        {
                            if (bindAttr.skipDuplicates)
                            {
                                // skip
                                continue;
                            }
                            else
                            {
                                throw new Exception($"[ExcelLoader] Duplicate key {convertedKey} in dict field={field.Name}");
                            }
                        }
                        dictID[convertedKey] = kvp.Value;
                    }
                }
            }
            else
            {
                Debug.LogWarning($"[ExcelLoader] field {field.Name}: dictionary ValueType != {dataType.Name}");
            }
        }
        // (b) 단일
        else if (field.FieldType == dataType)
        {
            if (dataList.Count > 0)
                field.SetValue(container, dataList[0]);
        }
        // (c) 리스트
        else if (IsListOfType(field.FieldType, dataType))
        {
            var listVal = field.GetValue(container) as System.Collections.IList;
            if (listVal == null)
            {
                var newList = Activator.CreateInstance(field.FieldType);
                field.SetValue(container, newList);
                listVal = newList as System.Collections.IList;
            }
            foreach (var obj in dataList)
            {
                listVal.Add(obj);
            }
        }
        else
        {
            Debug.LogWarning($"[ExcelLoader] field {field.Name} has [SheetBinding({bindAttr.SheetName})], but type mismatch? {field.FieldType}");
        }
    }

    private static bool IsListOfType(Type t, Type elem)
    {
        if (!t.IsGenericType) return false;
        if (t.GetGenericTypeDefinition() != typeof(List<>)) return false;
        return t.GetGenericArguments()[0] == elem;
    }
    /// <summary>
    /// Dictionary<K, V> 인지 판별, K/V 타입을 out으로 얻는다.
    /// 예) Dictionary<int, UnitData> → keyType=int, valType=UnitData
    /// </summary>
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

    /// <summary>
    /// string → keyType으로 변환
    /// 예) keyType == int → int.TryParse
    ///     keyType == enum → Enum.TryParse
    ///     keyType == string → 그대로
    /// 실패 시 null 반환
    /// </summary>
    private static object ConvertKeyString(string strKey, Type keyType)
    {
        if (keyType == typeof(string))
        {
            // 그대로
            return strKey;
        }
        else if (keyType == typeof(int))
        {
            if (int.TryParse(strKey, out int iVal))
                return iVal;
            return null;
        }
        else if (keyType == typeof(long))
        {
            if (long.TryParse(strKey, out long lVal))
                return lVal;
            return null;
        }
        else if (keyType.IsEnum)
        {
            if (Enum.TryParse(keyType, strKey, true, out object enVal))
                return enVal;
            return null;
        }
        else
        {
            // float, double 등 더 추가 가능
            try
            {
                return Convert.ChangeType(strKey, keyType);
            }
            catch
            {
                return null;
            }
        }
    }
}
