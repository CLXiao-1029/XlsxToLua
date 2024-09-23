using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using XlsxToLua.Common;

namespace XlsxToLua.Data;

internal static class MultiLanguageHelper
{
    private const string Filename = "TranslateMain.xlsx";
    /// <summary>
    /// 多语言key缓存值
    /// </summary>
    private static readonly Dictionary<object, object> KeyCacheData = new();
    /// <summary>
    ///多语言数据列表
    /// </summary>
    private static readonly Dictionary<object, List<object>> MultilingualList = new();
    /// <summary>
    /// 语种标签列表
    /// </summary>
    private static readonly List<LanguageTag> MultiLanguageTagList = new();
    private static readonly Dictionary<object, List<object>> MultilingualListTmp = new();
    private static bool _isNewFile = false;
    
    private static string? _filePath = default;
    private static string FilePath
    {
        get
        {
            if (_filePath != default) return _filePath;
            if (MainArgs.ConfigPath != string.Empty)
            {
                _filePath = Path.Combine(MainArgs.ConfigPath, Filename);
            }
            else
            {
                throw new Exception("ConfigPath is null");
            }

            return _filePath;
        }
    }

    private static string? _luaName;

    private static string LuaName
    {
        get
        {
            if (MainArgs.ExportPath != string.Empty)
            {
                var name = MainArgs.LuaFileName("i18n");
                _luaName = Path.Combine(MainArgs.ExportPath, name);
            }
            else
            {
                throw new Exception("ExportPath is null");
            }

            return _luaName; //Path.Combine(MainArgs.ExportPath, "i18n.lua");
        }
    }
    
    /// <summary>
    /// 加载多语言文件
    /// </summary>
    internal static void Load()
    {
        KeyCacheData.Clear();
        MultilingualList.Clear();
        MultiLanguageTagList.Clear();
        MultilingualListTmp.Clear();

        if (!File.Exists(FilePath)) return;
        
        if (!File.Exists(LuaName))
        {
            _isNewFile = true;
        }
        
        // 读取多语言文件

        using var fileSteam = File.Open(FilePath,FileMode.Open,FileAccess.Read,FileShare.ReadWrite);
        var package = new ExcelPackage(fileSteam);
        // 获取第一个工作表
        var worksheet = package.Workbook.Worksheets.First();
        Logger.Info($"成功加载 {Filename} 表的【{worksheet.Name}】工作簿");
        
        // 获取有效行数和列数
        var rowCount = worksheet.Dimension.Rows;
        var columnCount = worksheet.Dimension.Columns;
        
        // 缓存多语言key
        for (var row = 4; row <= rowCount; ++row)
        {
            var cell = worksheet.Cells[row, 1];
            var mainKey = cell.Value;
            if (mainKey == null) continue;
            var cache = $"{row}:{cell.Address}";
            if (!KeyCacheData.TryAdd(mainKey, cache))
            {
                Logger.Error($"重复多语言Key[{mainKey}]！当前 {cache} 行与之前 {KeyCacheData[mainKey]} 行的[{mainKey}]重复。");
            }
        }
            
        // 获取注解，第一列是多语言Key，从第二列开始
        for (var column = 2; column <= columnCount; ++column)
        {
            // 获取前两行
            var annotation = worksheet.Cells[1, column];
            var value = worksheet.Cells[2, column];
            if (value.Value == null) continue;
            MultiLanguageTagList.Add(new LanguageTag(value.Value, annotation.Value));
        }

        // 获取多语言数据，从第四行开始，前三行作为注解使用
        for (var row = 4; row <= rowCount; ++row)
        {
            // 获取多语言key
            var cell = worksheet.Cells[row, 1];
            var mainKey = cell.Value;
            var cache = $"{row}:{cell.Address}";
            if (mainKey != null)
            {
                var values = new List<object>();
                for (var column = 2; column <= columnCount; ++column)
                {
                    var value = worksheet.Cells[row, column].Value;
                    values.Add(value);
                }

                if (!MultilingualList.TryAdd(mainKey, values))
                {
                    Logger.Error($"重复多语言Key[{mainKey}]！当前 {cache} 行与之前 {KeyCacheData[mainKey]} 行的[{mainKey}]重复。");
                }
            }
            else
            {
                Logger.Warning($"{cache} 多语言Key为空");
            }
        }
        
        // 获取附表
        if (package.Workbook.Worksheets.Count <= 1) return;
        var worksheet1 = package.Workbook.Worksheets[1];
        if (worksheet1 == null) return;
        for (var row = 4; row <= rowCount; ++row)
        {
            var mainKey = worksheet1.Cells[row, 1].Value;
            if (mainKey == null) return;
            var values = new List<object>();
            for (var column = 2; column <= columnCount; ++column)
            {
                var value = worksheet1.Cells[row, column].Value;
                values.Add(value);
            }

            MultilingualListTmp.Add((string)mainKey, values);
        }
    }

    internal static StringBuilder LoadComment()
    {
        // 写出i18n
        var i18NComment = new StringBuilder();
        i18NComment.AppendLine("---@class Cfg_i18n");
        i18NComment.AppendLine($"---@field language string");
        foreach (var (key, _) in MultilingualList)
        {
            i18NComment.AppendLine($"---@field {key} string");
        }
        return i18NComment;
    }

    internal static void Save()
    {
        var fileInfo = new FileInfo(FilePath);
        if (fileInfo.Exists)
        {
            if (!_isNewFile) return;
            _isNewFile = false;
            Logger.Info($"{Filename} 文件已存在，准备覆盖文件。");
            fileInfo.Delete();
            ReplaceAndSave();
            I18N();
        }
        else
        {
            CreateAndSave(fileInfo);
            I18N();
        }
    }

    /// <summary>
    /// 尝试添加多语言数据
    /// </summary>
    /// <param name="key"></param>
    /// <param name="value"></param>
    internal static void TryAdd(object key, object? value)
    {
        if (value == null)
        {
            return;
        }

        if (MultilingualList.TryGetValue(key,out var values))
        {
            var oldValue = values[0];
            if (oldValue.Equals(value))
            {
                return;
            }
            
            _isNewFile = true;
            Logger.Error($"尝试添加多语言key[{key}]时，发现已存在相同 key[{key}]！New:{value}，Old:{oldValue}");
            if (MultilingualListTmp.ContainsKey(key))
            {
                // 刷新差异表数据
                MultilingualListTmp[key] = new List<object>() { oldValue, value };
            }
            else
            {
                // 添加差异表数据
                MultilingualListTmp.Add(key, new List<object>() { values[0], value });
            }
        }
        else
        {
            _isNewFile = true;
            MultilingualList.Add(key, new List<object>() { value });
        }
    }

    private static void I18N()
    {
        var dic = new Dictionary<object, StringBuilder>();
        
        // 写出i18n
        var i18N = new StringBuilder();
        i18N.AppendLine("---@type Cfg_i18n");
        i18N.AppendLine("local cfg_i18n = {");
        i18N.AppendLineIndent("language = \"en\"");
        i18N.AppendLine("}");
        i18N.AppendLine();
        
        //添加原表操作
        i18N.AppendLine("setmetatable(cfg_i18n, {");
        i18N.AppendLineIndent("__index = function (t, key)");
        ++Values.IndentLevel;
        i18N.AppendLineIndent("local languages = require (\"Data.\" .. t.language)");
        i18N.AppendLineIndent("if languages[key] == nil then");
        ++Values.IndentLevel;
        i18N.AppendLineIndent("CS.UnityEngine.Debug.LogError((\"多语言表的[%s]语种中不存在key[%s] languages[key] == nil\\n%s\"):format(tostring(t.language),tostring(key), debug.traceback()))");
        --Values.IndentLevel;
        i18N.AppendLineIndent("end");
        i18N.AppendLineIndent("return languages[key]");
        --Values.IndentLevel;
        i18N.AppendLineIndent("end");
        i18N.AppendLine("})");
        i18N.AppendLine();
        i18N.AppendLine("return cfg_i18n");
        dic.Add("i18n", i18N);
        
        for (var i = 0; i < MultiLanguageTagList.Count; i++)
        {
            var tagKey = MultiLanguageTagList[i].Value;
            var stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("---@type Cfg_i18n");
            stringBuilder.AppendLine("local language = {");
            foreach (var (key, values) in MultilingualList)
            {
                if (i >= values.Count) continue;
                var value = values[i];
                stringBuilder.AppendLineIndent($"{key} = \"{value}\",");
            }
            stringBuilder.RemoveTrailingComma();
            // stringBuilder.Remove(stringBuilder.Length - 1, 1);
            stringBuilder.AppendLine("}");
            stringBuilder.AppendLine();
            stringBuilder.AppendLine("return language");
            dic.Add(tagKey, stringBuilder);
        }
        
        foreach (var (key, stringBuilder) in dic)
        {
            if (MainArgs.ExportPath == null) continue;
            var name = MainArgs.LuaFileName(key);
            var savePath = Path.Combine(MainArgs.ExportPath, name);
            if (File.Exists(savePath))
                File.Delete(savePath);

            Logger.Info($"写出 lua 文件：{name}");
            File.WriteAllText(savePath, stringBuilder.ToString());
        }
    }
    
    private static void ReplaceAndSave()
    {
        CreateAndSave(new FileInfo(FilePath));
    }

    private static void CreateAndSave(FileInfo fileInfo)
    {
        using var package = new ExcelPackage(fileInfo);
        if (package.Workbook.Worksheets.Count == 0)
            package.Workbook.Worksheets.Add("翻译表");

        var worksheet = package.Workbook.Worksheets.First();

        // 添加表头
        worksheet.Cells[1, 1].Value = "多语言Key";
        worksheet.Cells[2, 1].Value = "Key";
        for (var i = 0; i < MultiLanguageTagList.Count; i++)
        {
            var tag = MultiLanguageTagList[i];
            worksheet.Cells[1, i + 2].Value = tag.Name;
            worksheet.Cells[2, i + 2].Value = tag.Value;
        }

        Logger.Info($"准备写出【{worksheet.Name}】");

        // 写出数据
        var keys = MultilingualList.Keys.ToArray();
        for (var i = 0; i < keys.Length; i++)
        {
            var key = keys[i];
            worksheet.Cells[i + 4, 1].Value = key;
            var values = MultilingualList[key];
            for (var j = 0; j < values.Count; j++)
            {
                worksheet.Cells[i + 4, j + 2].Value = values[j];
            }
        }

        // 修改表头背景颜色
        for (var i = 0; i < 3; i++)
        {
            AddHeader(worksheet.Rows[i + 1]);
        }

        // 写出附表
        if (MultilingualListTmp.Count <= 0)
        {
            // 保存
            package.Save();
            return;
        }

        package.Workbook.Worksheets.Add("差异表");
        var worksheet1 = package.Workbook.Worksheets[1];
        worksheet1.Cells[1, 1].Value = "多语言Key";
        worksheet1.Cells[2, 1].Value = "key";
        worksheet1.Cells[1, 2].Value = "现在的文本";
        worksheet1.Cells[2, 2].Value = "zh";
        worksheet1.Cells[1, 3].Value = "之前的文本";
        worksheet1.Cells[2, 3].Value = "zh";

        Logger.Info($"准备写出【{worksheet.Name}】");
        // 添加内容
        var tmpKeys = MultilingualListTmp.Keys.ToArray();
        for (var i = 0; i < tmpKeys.Length; i++)
        {
            var key = tmpKeys[i];
            var values = MultilingualListTmp[key];
            //从第四行第一列开始
            worksheet1.Cells[i + 4, 1].Value = key;
            for (var j = 0; j < values.Count; j++)
            {
                //第二列开始
                worksheet1.Cells[i + 4, j + 2].Value = values[j];
            }
        }

        for (var i = 0; i < 3; i++)
        {
            AddHeader(worksheet.Rows[i + 1]);
        }

        // 保存
        package.Save();
    }
    private static void AddHeader(ExcelRangeRow row)
    {
        row.Style.Fill.PatternType = ExcelFillStyle.Solid;
        row.Style.Fill.BackgroundColor.SetColor(255, 112, 173, 71);
    }
}