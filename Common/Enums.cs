namespace XlsxToLua.Common;
internal enum TortoiseType
{
    None,
    Git,
    Svn
}

/// <summary>
/// 命名规则
/// </summary>
internal enum NameRule
{
    None,
    /// <summary>
    /// 首字母大写
    /// </summary>
    CapitalizeFirst,
    /// <summary>
    /// 首字母小写
    /// </summary>
    LowercaseFirst,
    /// <summary>
    /// 转大写
    /// </summary>
    UpperCase,
    /// <summary>
    /// 转小写
    /// </summary>
    LowerCase,
    /// <summary>
    /// 驼峰命名
    /// </summary>
    CamelCase,
    /// <summary>
    /// 单词连接首字母大写
    /// </summary>
    PascalCase,
}