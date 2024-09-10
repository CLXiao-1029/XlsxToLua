using OfficeOpenXml;

namespace XlsxToLua.Data;

internal struct LanguageTag
{
    public object? Name;
    public object Value;

    public LanguageTag(object value, object? name)
    {
        Value = value;
        Name = name;
    }
}

/// <summary>
/// 一张表的数据结构
/// </summary>
internal struct TableDataInfo
{
    /// <summary>
    /// 行数
    /// </summary>
    public readonly int Rows;

    /// <summary>
    /// 列数
    /// </summary>
    public readonly int Columns;

    /// <summary>
    /// 文件名
    /// </summary>
    public readonly string FileName;

    /// <summary>
    /// 表名
    /// </summary>
    public readonly string TableName;

    /// <summary>
    /// 所属文件夹
    /// </summary>
    public readonly string? FolderName;

    /// <summary>
    /// 表数据
    /// </summary>
    public readonly ExcelRange Cells;

    public TableDataInfo(int rows, int columns, string fileName, string tableName, string? folderName, ExcelRange cells)
    {
        Rows = rows;
        Columns = columns;
        FileName = fileName;
        TableName = tableName;
        FolderName = folderName;
        Cells = cells;
    }
}