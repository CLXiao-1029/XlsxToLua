using System.Text;
using OfficeOpenXml;
using XlsxToLua.Common;

namespace XlsxToLua.Data;

public class LuaExportHelper
{
    private static readonly Dictionary<string, string> Comment = new ();
    private static readonly Dictionary<string, StringBuilder> Content = new ();
    
    private static readonly Dictionary<object, string> KeyCache = new ();
    private static readonly Dictionary<object, string> ValueCache = new ();

    #region Query All
    
    /// <summary>
    /// 查询Excel所有数据
    /// </summary>
    public static void QueryXlsxAll()
    {
        if (MainArgs.ConfigPath == string.Empty) 
            return;
        // 区分模式
        if (MainArgs.Summary)
        {
            var fileInfoDic = new Dictionary<string, List<FileInfo>>();
            FileUtils.GetTopDirectoryFiles(MainArgs.ConfigPath, ref fileInfoDic);
            foreach (var (folder, fileInfos) in fileInfoDic)
            {
                ForeachXlsxData(fileInfos, folder);
            }
        }
        else
        {
            var fileInfos = new List<FileInfo>();
            FileUtils.GetDirectoryFiles(MainArgs.ConfigPath, ref fileInfos);
            ForeachXlsxData(fileInfos);
        }
    }

    private static void ForeachXlsxData(IEnumerable<FileInfo> fileInfos, string? folder = null)
    {
        foreach (var fileInfo in from fileInfo in fileInfos
                 where !fileInfo.Name.StartsWith("~$")
                 let ext = fileInfo.Extension.ToLower()
                 where ext.Equals(".xlsx")
                 select fileInfo)
        {
            ReadXlsxData(fileInfo,folder);
        }
    }

    private static void ReadXlsxData(FileInfo fileInfo, string? folder)
    {
        //检查文件名
        if (fileInfo.Name.StartsWith("~$") || fileInfo.Name.StartsWith("Translate"))
        {
            return;
        }

        //检查扩展名
        var ext = fileInfo.Extension.ToLower();
        if (!ext.Equals(".xlsx"))
        {
            return;
        }
        
        try
        {
            using var fileStream = fileInfo.Open(FileMode.Open,FileAccess.Read,FileShare.ReadWrite);
            var package = new ExcelPackage(fileStream);
            foreach (var sheet in package.Workbook.Worksheets)
            {
                var sheetNames = sheet.Name.Split('|');
                if (sheetNames.Length < 2) continue;
                    
                // 获取表名,行数,列数
                var tableName = sheetNames.Last();
                var rowCount = sheet.Dimension.Rows;
                var columnCount = sheet.Dimension.Columns;
                    
                // 检测有效行数
                if (rowCount < MainArgs.StartRow) continue;
                    
                // 构建数据
                var data = new TableDataInfo(rowCount, columnCount, fileInfo.Name, tableName, folder, sheet.Cells);
                Logger.Info($"加载文件{fileInfo.Name}中的{tableName}工作簿.");
                // 准备加入数据
                var key = tableName;
                if (MainArgs.Summary)
                    key = $"{folder}|{tableName}";

                if (Values.DataTables.TryGetValue(key, out var existInfo))
                {
                    const string errFormat = "导出配置表文件【{0}】的工作簿“{1}”时，发现已经被配置表文件【{2}】占用，请修改当前配置表工作簿名“{3}”";
                    var error = string.Format(errFormat, data.FileName, data.TableName, existInfo.FileName, data.TableName);
                    Logger.ErrorAndExit(error);
                    continue;
                }
                    
                Values.DataTables.Add(key, data);
            }
            fileStream.Close();
            fileStream.Dispose();
        }
        catch (Exception e)
        {
            Logger.Exception(e);
        }
    }
    
    #endregion

    public static void XlsxToLua()
    {
        var keys = Values.DataTables.Keys.ToArray();
        foreach (var key in keys)
        {
            var tableDateTime = DateTime.Now;
            Values.Reset();
            
            Logger.Info($"============准备导表 {key}============");
            var tableDataInfo = Values.DataTables[key];
            Logger.Info($"{key} 开始解析字段类型");
            FillAnnotation(tableDataInfo.Cells, tableDataInfo.Columns);
            Logger.Info($"{key} 开始解析lua表");
            if (MainArgs.Summary)
            {
                AddComment(tableDataInfo.TableName, GetAnnotation(tableDataInfo.TableName));
                if (tableDataInfo.FolderName == null)
                {
                    Logger.Error($"{key} 导表失败，请检查文件夹路径是否正确");
                }
                else
                {
                    AddLuaContent(tableDataInfo.FolderName, XlsxToLuaContent(tableDataInfo));
                }
            }
            else
            {
                if (MainArgs.SplitComment)
                    AddComment(tableDataInfo.TableName, GetAnnotation(tableDataInfo.TableName));
                
                var content = XlsxToLuaContent(tableDataInfo);
                var fileName = MainArgs.LuaFileName(tableDataInfo.TableName);
                FileUtils.SafeSave(fileName, content.ToString());
            }

            Logger.Info(MainArgs.ShowTime
                ? $"{tableDataInfo.FileName} - {tableDataInfo.TableName} 导表完成 耗时：{(DateTime.Now - tableDateTime).TotalMilliseconds}"
                : $"{tableDataInfo.FileName} - {tableDataInfo.TableName} 导表完成");
        }
        if (MainArgs.Summary)
        {
            Logger.Info("开始生成lua总表...");
            var all = SummaryTable();
            var allKeys = all.Keys.ToArray();
            foreach (var name in allKeys)
            {
                var content = all[name];
                var fileName = MainArgs.LuaFileName(name);
                if (MainArgs.ExportPath != null)
                {
                    FileUtils.SafeSave(fileName, content.ToString());
                }
            }
        }

        if (!MainArgs.SplitComment) return;
        Logger.Info("生成lua注释 . . .");
        FileUtils.SafeSave(MainArgs.LuaFileName("ConfigComment"), CommentToString());
    }
    
    /// <summary>
    /// 填充注解信息
    /// </summary>
    /// <param name="cells"></param>
    /// <param name="columns"></param>
    private static void FillAnnotation(ExcelRange cells, int columns)
    {
        for (var j = 1; j <= columns; ++j)
        {
            var info = cells[1, j].Value;
            var name = cells[2, j].Value;
            var type = cells[3, j].Value;
            if (name != null && type != null)
            {
                var cache = $"{j}:{cells.Address}";
                if (!ValueCache.TryAdd(name, cache))
                {
                    Logger.Error($"{cells.Worksheet.Name} 表内存在重复的字段名[{name}]！当前 {cache} 列与之前 {ValueCache[name]} 列的定义重复");
                }
                Values.VarInfos.Add(info.ToString());
                Values.VarNames.Add(name.ToString());
                Values.VarTypes.Add(type.ToString());
            }
            else
            {
                Values.VarInfos.Add(null);
                Values.VarNames.Add(null);
                Values.VarTypes.Add(null);
            }
        }
    }
    
    /// <summary>
    /// 获取注解文本
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    private static string GetAnnotation(string name)
    {
        var content = new StringBuilder();
        GenAnnotation(name, content);
        return content.ToString();
    }
    
    /// <summary>
    /// 生成注解文本
    /// </summary>
    /// <param name="name"></param>
    /// <param name="content"></param>
    private static void GenAnnotation(string name, StringBuilder content)
    {
        content.AppendLine($"---@class Cfg_{name}");
        for (var i = 0; i < Values.VarTypes.Count; i++)
        {
            var type = Values.VarTypes[i];
            if (Values.VarTypeIsNull(i))
            {
                continue;
            }

            type = Values.AnalyzeType(type);
            if (Values.VarTypeDesc.Length != 0)
            {
                content.AppendLine($"---@field {Values.VarNames[i]} {type} {Values.VarInfos[i]} {Values.VarTypeDesc}");
            }
            else
            {
                content.AppendLine($"---@field {Values.VarNames[i]} {type} {Values.VarInfos[i]}");
            }
        }

        content.AppendLine();
    }
    
    
    /// <summary>
    /// 添加注释
    /// </summary>
    /// <param name="name">表名</param>
    /// <param name="content">注释内容</param>
    /// <returns></returns>
    private static bool AddComment(string name, string content)
    {
        return Comment.TryAdd(name, content);
    }
    
    /// <summary>
    /// 注释转文本
    /// </summary>
    /// <returns></returns>
    private static string CommentToString()
    {
        var comment = new StringBuilder();
        foreach (var (key, value) in Comment)
        {
            comment.Append(value);
        }

        comment.Append('\n');
        if (MainArgs.Tortoise != TortoiseType.None)
            comment.AppendLine($"--ConfigLatestCommit:{TortoiseHelper.CommitLog}");

        return comment.ToString();
    }

    /// <summary>
    /// 添加Lua内容
    /// </summary>
    /// <param name="name">表名</param>
    /// <param name="content">内容</param>
    /// <returns></returns>
    private static void AddLuaContent(string name, StringBuilder content)
    {
        if (Content.TryGetValue(name, out var stringBuilder))
        {
            stringBuilder.AppendLine(content.ToString());
        }
        else
        {
            content.AppendLine();
            Content.Add(name, content);
        }
    }

    /// <summary>
    /// 一张表转换为lua格式文本
    /// </summary>
    /// <param name="dataInfo"></param>
    /// <returns></returns>
    private static StringBuilder XlsxToLuaContent(TableDataInfo dataInfo)
    {
        var tableName = dataInfo.TableName;
        var columnCount = dataInfo.Columns;
        var rowCount = dataInfo.Rows;
        var content = new StringBuilder();
        var isSummary = MainArgs.Summary;

        if (isSummary)
        {
            content.AppendLineIndent($"---@type Cfg_{tableName}[]");
            tableName = MainArgs.Name(tableName);
            content.AppendLineIndent($"{tableName} = {{");
        }
        else
        {
            if (MainArgs.Tortoise != TortoiseType.None)
            {
                content.AppendLine($"--ConfigLatestCommit:{TortoiseHelper.CommitLog}");
            }

            if (MainArgs.SplitComment)
                content.AppendLine($"---@type Cfg_{tableName}[]");
            else
                GenAnnotation(tableName, content);

            content.AppendLine($"local {tableName} = {{");
            --Values.IndentLevel;
        }
        
        KeyCache.Clear();
        ValueCache.Clear();
        
        var starIdx = MainArgs.StartRow;
        for (var row = starIdx; row <= rowCount; ++row)
        {
            var cell = dataInfo.Cells[row, 1];
            var valueId = cell.Value;
            var cache = $"{row}:{cell.Address}";
            if (valueId == null)
            {
                Logger.Warning($"{cell.Worksheet.Name} 表内存在空行！当前 {cache} 行为空行");
                continue;
            }
            
            // 忽略重复的ID
            if (!KeyCache.TryAdd(valueId, cache))
            {
                Logger.Error($"{cell.Worksheet.Name} 表内存在重复ID[{valueId}]！当前 {cache} 行与之前 {KeyCache[valueId]} 行的id重复");
                continue;
            }
                
            ++Values.IndentLevel;
            if (MainArgs.Order)
            {
                content.AppendLineIndent('{');
            }
            else
            {
                content.AppendLineIndent($"[{valueId}] = {{");
            }
            
            for (var column = 1; column < columnCount + 1; ++column)
            {
                var columnIndex = column - 1;
                if (Values.VarTypeIsNull(columnIndex))
                {
                    continue;
                }

                var varKey = Values.VarNames[columnIndex];
                var varType = Values.VarTypes[columnIndex];

                var rowData = dataInfo.Cells[row, column].Value;
                var isNull = rowData == null;
                if (!isNull)
                {
                    if (ExcelErrorValue.Values.IsErrorValue(rowData))
                    {
                        Logger.Error($"表名：{tableName} ,第{row}行{column}列，key = {varKey}，单元格值错误：{rowData}");
                        isNull = true;
                    }
                }

                var isTranslationVal = false;
                if (MainArgs.ExtractText)
                {
                    var isTranslation = dataInfo.Cells[4, column].Value?.ToString();
                    if (!string.IsNullOrEmpty(isTranslation) && !string.IsNullOrWhiteSpace(isTranslation))
                    {
                        isTranslationVal = isTranslation.Equals("true", StringComparison.OrdinalIgnoreCase);
                    }
                }

                if (isTranslationVal)
                {
                    var keyId = $"{tableName}_{varKey}_{valueId}";
                    if (rowData != null)
                    {
                        MultiLanguageHelper.TryAdd(keyId, rowData);
                    }

                    rowData = keyId;
                }

                content.GenerateLuaLine(tableName,  varKey, varType, rowData, isNull, row, column);
            }

            content.AppendIndent("},");
            content.AppendLineIndent("");
            --Values.IndentLevel;
        }

        if (isSummary)
        {
            content.AppendLineIndent("},");
        }
        else
        {
            ++Values.IndentLevel;
            content.AppendLine("}");
            content.AppendLine();
            content.AppendLine($"return {tableName}");
        }

        return content;
    }
    
    /// <summary>
    /// 获取总表文本
    /// </summary>
    /// <returns></returns>
    private static Dictionary<string, StringBuilder> SummaryTable()
    {
        Dictionary<string, List<string>> folderTables = new();
        string? curFolderName = null;
        foreach (var item in Values.DataTables)
        {
            if (curFolderName != item.Value.FolderName)
            {
                curFolderName = item.Value.FolderName;
                if (curFolderName != null)
                    folderTables.Add(curFolderName, new List<string>());
            }

            if (curFolderName != null)
                folderTables[curFolderName].Add(item.Value.TableName);
        }

        var comment = new StringBuilder();
        var folders = folderTables.Keys.ToArray();
        foreach (var t in folders)
        {
            comment.Clear();

            comment.AppendLine($"---@class Cfg_{t}");
            foreach (string tableName in folderTables[t])
            {
                comment.AppendLine($"---@field {tableName} Cfg_{tableName}[]");
            }

            comment.AppendLine();
            AddComment(t, comment.ToString());
        }

        var builders = new Dictionary<string, StringBuilder>();
        foreach (var name in folders)
        {
            var merge = new StringBuilder();
            if (MainArgs.Tortoise != TortoiseType.None)
            {
                merge.AppendLine($"--ConfigLatestCommit:{TortoiseHelper.CommitLog}");
            }

            merge.AppendLine($"---@type Cfg_{name}");

            merge.AppendLine($"local cfg_{name} = {{");

            var content = Content[name];
            merge.Append($"{content}");
            merge.AppendLine("}");
            merge.AppendLine();

            merge.Append($"return cfg_{name}");
            builders.Add(name, merge);
        }

        return builders;
    }
}