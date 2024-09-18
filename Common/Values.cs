using System.Text;
using System.Text.RegularExpressions;
using XlsxToLua.Data;

namespace XlsxToLua.Common;

internal static class Values
{
    private const string TablePatten = @"(?i)(?<=\[)(.*)(?=\])";
    
    private const string LuaTableIndentationString = "\t";
    
    /// <summary>
    /// 配置表数据缓存
    /// </summary>
    internal static Dictionary<string, TableDataInfo> DataTables = new();
    
    /// <summary>
    /// 数据列描述
    /// </summary>
    internal static List<string?> VarInfos = new();
    /// <summary>
    /// 数据列字段名字
    /// </summary>
    internal static List<string?> VarNames = new();
    /// <summary>
    /// 数据列字段类型
    /// </summary>
    internal static List<string?> VarTypes = new();
    /// <summary>
    /// 数据类型描述
    /// </summary>
    internal static StringBuilder VarTypeDesc = new();

    /// <summary>
    /// 缩进值
    /// </summary>
    internal static int IndentLevel = 1;
    
    public static string IndentIndex()
    {
        var stringBuilder = new StringBuilder();
        for (var i = 0; i < IndentLevel; ++i)
            stringBuilder.Append(LuaTableIndentationString);

        return stringBuilder.ToString();
    }
    
    public static void Reset()
    {
        IndentLevel = 1;
        VarInfos.Clear();
        VarNames.Clear();
        VarTypes.Clear();
    }
    
    public static bool VarTypeIsNull(int index)
    {
        return string.IsNullOrEmpty(VarTypes[index]) || string.IsNullOrWhiteSpace(VarTypes[index]);
    }

    #region Check Data

    private static bool EqualsOrdinalIgnoreCase(this string str, string? value, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
    {
        return str.Equals(value, comparisonType);
    }
    
    /// <summary>
    /// 检测数据类型是否为数字
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    private static bool IsNumber(string type)
    {
        return type.EqualsOrdinalIgnoreCase("int") || type.EqualsOrdinalIgnoreCase("float") || type.EqualsOrdinalIgnoreCase("double") || type.EqualsOrdinalIgnoreCase("long");
    }

    /// <summary>
    /// 检测数据类型是否为字符串
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    private static bool IsString(string type)
    {
        return type.EqualsOrdinalIgnoreCase("string");
    }

    private static bool IsBool(string type)
    {
        return type.EqualsOrdinalIgnoreCase("bool") || type.EqualsOrdinalIgnoreCase("boolean");
    }

    private static bool IsTable(string type)
    {
        return type.EqualsOrdinalIgnoreCase("table");
    }

    private static bool IsNumberArray(string type)
    {
        var type1 = type.EqualsOrdinalIgnoreCase("int[]") || type.EqualsOrdinalIgnoreCase("float[]") || type.EqualsOrdinalIgnoreCase("double[]") || type.EqualsOrdinalIgnoreCase("long[]");
        var type2 = type.EqualsOrdinalIgnoreCase("array[int]") || type.EqualsOrdinalIgnoreCase("array[float]") || type.EqualsOrdinalIgnoreCase("array[double]") || type.EqualsOrdinalIgnoreCase("array[long]");
        return type1 || type2;
    }

    private static bool IsStringArray(string type)
    {
        return type.EqualsOrdinalIgnoreCase("string[]") || type.EqualsOrdinalIgnoreCase("array[string]");
    }

    private static bool IsBoolArray(string type)
    {
        return type.EqualsOrdinalIgnoreCase("bool[]") || type.EqualsOrdinalIgnoreCase("boolean[]") || type.EqualsOrdinalIgnoreCase("array[bool]") || type.EqualsOrdinalIgnoreCase("array[boolean]");
    }

    private static bool IsTableArray(string type)
    {
        return type.EqualsOrdinalIgnoreCase("table[]") || type.EqualsOrdinalIgnoreCase("array[table]");
    }

    private static bool IsArrayTable(string type)
    {
        var match = Regex.Match(type, Values.TablePatten);

        return match.Success;
    }

    #endregion

    #region Convert Data

    
    private static bool ToBool(string? data, out string? error)
    {
        error = null;
        data = data?.ToLower();
        var boolean = false;
        switch (data)
        {
            case "1":
            case "true":
                boolean = true;
                break;
            case "0":
            case "false":
                boolean = false;
                break;
            default:
                error = $"{data} is not a boolean.非法的布尔值";
                break;
        }

        return boolean;
    }

    /// <summary>
    /// 类型转换默认值
    /// </summary>
    /// <param name="type"></param>
    /// <param name="data"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    private static object? ParseDefault(string type, string data, out string? error)
    {
        type = type.ToLower();
        error = null;
        object? result = null;
        switch (type)
        {
            case "int":
                if (int.TryParse(data, out var intResult))
                {
                    result = intResult;
                }
                else
                {
                    error = $"{data} is not a int.非法的整数";
                }

                break;
            case "float":
                if (float.TryParse(data, out var floatResult))
                {
                    result = floatResult;
                }
                else
                {
                    error = $"{data} is not a float.非法的小数";
                }

                break;
            case "double":
                if (double.TryParse(data, out var doubleResult))
                {
                    result = doubleResult;
                }
                else
                {
                    error = $"{data} is not a double.非法的小数";
                }

                break;
            case "long":
                if (long.TryParse(data, out var longResult))
                {
                    result = longResult;
                }
                else
                {
                    error = $"{data} is not a long.非法的大整数";
                }

                break;
            case "bool":
            case "boolean":
                var boolean = ToBool(data, out error);
                result = boolean.ToString().ToLower();
                break;
            case "string":
                result = string.IsNullOrEmpty(data) ? null : $"\"{data}\"";
                break;
            case "table":
                result = string.IsNullOrEmpty(data) ? null : data;
                break;
            default:
                error = $"{type} is not a type.非法的类型";
                break;
        }

        return result;
    }

    /// <summary>
    /// 分析表数据
    /// </summary>
    /// <param name="type"></param>
    /// <param name="data"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    private static List<Dictionary<string, object?>>? AnalyzeTable(string type, string? data, out string? error)
    {
        error = null;
        try
        {
            var match = Regex.Match(type, TablePatten);
            // 拆分数据类型结构
            var keyValueTypes = match.Value.Split(',');
            // 拆分数据组
            var valueGroups = data?.Split('|');
            var children = new List<Dictionary<string, object?>>();
            if (valueGroups != null)
            {
                foreach (var group in valueGroups)
                {
                    error = null;
                    if (string.IsNullOrEmpty(group)) continue;

                    var node = new Dictionary<string, object?>();
                    var values = group.Split(',');
                    // 检查数据格式有效长度
                    var length = values.Length > keyValueTypes.Length ? keyValueTypes.Length : values.Length;
                    for (var i = 0; i < length; i++)
                    {
                        var keyValues = keyValueTypes[i].Split(':');
                        var key = keyValues[0];
                        var valType = keyValues[1];
                        var value = values[i];

                        if (!string.IsNullOrEmpty(error)) break;
                        node.Add(key, ParseDefault(valType, value, out error));
                    }

                    if (!string.IsNullOrEmpty(error))
                    {
                        error += $"子表达式：{values}\n";
                        break;
                    }

                    children.Add(node);
                }
            }
            else
            {
                error = "AnalyzeTable data is null";
            }

            return string.IsNullOrEmpty(error) ? children : null;
        }
        catch (Exception e)
        {
            error += e;
            return null;
        }
    }

    /// <summary>
    /// 分析数组数据
    /// </summary>
    /// <param name="type"></param>
    /// <param name="data"></param>
    /// <param name="error"></param>
    /// <returns></returns>
    private static List<object?>? AnalyzeArray(string type, string? data, out string? error)
    {
        error = null;
        try
        {
            var typeName = type.Split('[');
            if (typeName[^1].Equals("]"))
            {
                type = typeName[0];
            }
            else
            {
                var match = Regex.Match(type, TablePatten);
                type = match.Value;
            }

            var children = new List<object?>();
            var values = data?.Split('|');
            if (values != null)
            {
                foreach (var value in values)
                {
                    if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(error))
                    {
                        children.Add(ParseDefault(type, value, out error));
                    }
                    else
                    {
                        error += $"子表达式：{value}\n";
                    }
                }
            }
            else
            {
                error = "AnalyzeArray data is null.";
            }

            return string.IsNullOrEmpty(error) ? children : null;
        }
        catch (Exception e)
        {
            error += e;
            return null;
        }
    }

    /// <summary>
    /// 分析数据类型并转换
    /// </summary>
    /// <param name="type"></param>
    /// <param name="isTop"></param>
    /// <returns></returns>
    public static string AnalyzeType(string? type, bool isTop = true)
    {
        if (isTop)
        {
            VarTypeDesc.Clear();
        }

        if (string.IsNullOrEmpty(type))
        {
            return "any";
        }

        var typeString = type.Trim();
        if (IsNumber(typeString))
        {
            return "number";
        }

        if (IsString(typeString) || IsTable(typeString))
        {
            return typeString;
        }

        if (IsBool(typeString))
        {
            return "boolean";
        }

        if (IsNumberArray(typeString))
        {
            return "number[]";
        }

        if (IsStringArray(typeString))
        {
            return "string[]";
        }

        if (IsTableArray(typeString))
        {
            return "table[]";
        }

        if (IsBoolArray(typeString))
        {
            return "boolean[]";
        }

        if (typeString.StartsWith("arrayTable", StringComparison.OrdinalIgnoreCase))
        {
            var arrayTable = new StringBuilder();
            arrayTable.Append("table<");

            if (isTop)
                VarTypeDesc.Append('{');

            var formatMatch = Regex.Matches(typeString, TablePatten);
            var matchValue = formatMatch[0].Value;
            var objects = matchValue.Split(new[] { ':', ',' });
            for (var i = 0; i < objects.Length; i++)
            {
                var index = i + 1;
                if (index % 2 == 0)
                {
                    arrayTable.Append(AnalyzeType(objects[i], false));
                    arrayTable.Append(',');
                }
                else
                {
                    if (isTop)
                    {
                        VarTypeDesc.Append(objects[i]);
                        VarTypeDesc.Append(',');
                    }
                }
            }

            arrayTable.Remove(arrayTable.Length - 1, 1);
            arrayTable.Append('>');

            if (!isTop) return arrayTable.ToString();
            VarTypeDesc.Remove(VarTypeDesc.Length - 1, 1);
            VarTypeDesc.Append('}');

            return arrayTable.ToString();
        }

        if (typeString.StartsWith("[") && typeString.EndsWith("]"))
        {
            var formatMatch = Regex.Match(typeString, TablePatten);
            if (formatMatch.Success)
            {
                var objs = formatMatch.Value.Split(new[] { ':', ',' });
                if (objs.Length == 0)
                {
                    var val = AnalyzeType(formatMatch.Value, false);
                    return $"{val}[]";
                }
            }

            var matchCollection = Regex.Matches(typeString, @"([a-z]+):([a-z]+)");
            var arrayTable = new StringBuilder();
            arrayTable.Append("table<");
            if (isTop)
            {
                VarTypeDesc.Append('{');
            }

            foreach (Match match in matchCollection)
            {
                var key = match.Groups[1].Value;
                var valType = match.Groups[2].Value;
                arrayTable.Append(AnalyzeType(valType, false));
                arrayTable.Append(',');

                if (!isTop) continue;
                VarTypeDesc.Append(key);
                VarTypeDesc.Append(',');
            }

            arrayTable.Remove(arrayTable.Length - 1, 1);
            arrayTable.Append('>');

            if (!isTop) return arrayTable.ToString();
            VarTypeDesc.Remove(VarTypeDesc.Length - 1, 1);
            VarTypeDesc.Append('}');

            return arrayTable.ToString();
        }

        return "any";
    }

    /// <summary>
    /// 生成Lua一行代码
    /// </summary>
    /// <param name="content"></param>
    /// <param name="name"></param>
    /// <param name="key"></param>
    /// <param name="type"></param>
    /// <param name="data"></param>
    /// <param name="isNull"></param>
    /// <param name="row"></param>
    /// <param name="column"></param>
    public static void GenerateLuaLine(this StringBuilder content, string name, string? key, string? type, object? data, bool isNull, int row = default,
        int column = default)
    {
        if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(key) || string.IsNullOrEmpty(type))
        {
            Logger.Error($"表名：{name} ,第{row}行{column}列，key = {key}，type = {type}");
            return;
        }
        var dataString = data?.ToString();
            
        ++IndentLevel;
        if (IsNumber(type))
        {
            if (!isNull)
            {
                if (int.TryParse(dataString, out _) ||
                    long.TryParse(dataString, out _) ||
                    float.TryParse(dataString, out _) ||
                    double.TryParse(dataString, out _))
                {
                    isNull = false;
                }
                else
                {
                    isNull = true;
                }

                if (!isNull)
                {
                    content.AppendLineIndent($"{key} = {data},");
                }
            }
        }
        else if (IsString(type))
        {
            if (!isNull)
            {
                if (!string.IsNullOrEmpty(dataString) && !string.IsNullOrWhiteSpace(dataString))
                {
                    content.AppendLineIndent($"{key} = \"{data}\",");
                }
            }
        }
        else if (IsBool(type))
        {
            if (!isNull)
            {
                var boolean = ToBool(dataString, out var error);
                if (error != null)
                    Logger.Error($"表名：{name} ,第{row}行{column}列，key = {key}，{error}");

                content.AppendLineIndent($"{key} = {boolean.ToString().ToLower()},");
            }
        }
        else if (IsTable(type))
        {
            if (!isNull)
                content.AppendLineIndent($"{key} = {data},");
        }
        else if (IsNumberArray(type)
                 || IsStringArray(type)
                 || IsTableArray(type)
                 || IsBoolArray(type))
        {
            if (!isNull)
            {
                content.AppendLineIndent($"{key} = {{");
                ++IndentLevel;
                var objects = AnalyzeArray(type, dataString, out var error);

                if (error != null)
                    Logger.Error($"表名：{name} ,第{row}行{column}列，{error}");

                if (objects != null)
                {
                    for (var i = 0; i < objects.Count; i++)
                    {
                        content.AppendLineIndent($"{objects[i]},");
                    }

                    content.Remove(content.Length - 1, 1);
                }

                --IndentLevel;
                content.AppendLineIndent("},");
            }
        }
        else if (IsArrayTable(type))
        {
            if (!isNull)
            {
                content.AppendLineIndent($"{key} = {{");
                ++IndentLevel;
                var dataObjs = AnalyzeTable(type, dataString, out var error);

                if (error != null)
                    Logger.Error($"表名：{name} ,第{row}行{column}列，{error}");

                if (dataObjs != null)
                {
                    foreach (var nodes in dataObjs)
                    {
                        content.AppendLineIndent("{");
                        ++IndentLevel;
                        foreach (var node in nodes)
                        {
                            content.AppendLineIndent($"{node.Key} = {node.Value},");
                        }

                        --IndentLevel;
                        content.Remove(content.Length - 1, 1);
                        content.AppendLineIndent("},");
                    }
                }

                --IndentLevel;
                content.AppendLineIndent("},");
            }
        }
        else if (type.Equals("any"))
        {
            if (!isNull)
                content.AppendLineIndent($"{key} = {data},");
        }
        else
        {
            Logger.Error($"表名：{name} ,第{row}行{column}列数据类型[{type}]错误，key = {key}");
        }

        --IndentLevel;
    }

    #endregion
}