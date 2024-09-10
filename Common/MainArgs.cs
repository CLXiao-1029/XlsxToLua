namespace XlsxToLua.Common;

internal static class MainArgs
{
    /// <summary>
    /// 配置表路径
    /// </summary>
    internal static string ConfigPath = string.Empty;
    
    /// <summary>
    /// 输出路径
    /// </summary>
    internal static string ExportPath = string.Empty;
    
    internal static int StartRow = 6;
    /// <summary>
    /// 扩展名
    /// </summary>
    private static string Extension = ".lua";
    
    /// <summary>
    /// 有序数据,有序状态下,数据为数组,否则为字典
    /// </summary>
    internal static bool Order = false;
    
    /// <summary>
    /// 总表模式。默认为false不开启。开启后会将所有表合并为一个总表，并按照配置表文件夹的根目录名字命名
    /// </summary>
    internal static bool Summary = false;
    
    /// <summary>
    /// 注解拆分，默认true开启。拆分后，注解会单独生成一个文件。
    /// </summary>
    internal static bool SplitComment = true;
    
    /// <summary>
    /// 文本提取，摘取表格中的文本并转换成翻译的key。默认true开启。
    /// </summary>
    internal static bool ExtractText = true;
    
    /// <summary>
    /// 显示时间流逝
    /// </summary>
    internal static bool ShowTime = false;
    
    /// <summary>
    /// 获取提交记录
    /// </summary>
    internal static TortoiseType Tortoise = TortoiseType.None;
    
    /// <summary>
    /// 命名规则
    /// </summary>
    internal static NameRule NamingRule = NameRule.None;

    internal static void ParseArgs(string[] args)
    {
        if (args.Length < 2)
        {
            Logger.ErrorAndExit("参数错误，请输入配置文件路径和导出路径");
            return;
        }
        
        ConfigPath = args[0];
        ExportPath = args[1];
        if (args.Length >= 2) if (!bool.TryParse(args[2], out Order)) Order = false;
        if (args.Length >= 3) if (!bool.TryParse(args[3], out Summary)) Summary = false;
        if (args.Length >= 4) if (!bool.TryParse(args[4], out SplitComment)) SplitComment = true;
        if (args.Length >= 5) if (!bool.TryParse(args[5], out ExtractText)) ExtractText = true;
        if (args.Length >= 6) if (!bool.TryParse(args[6], out ShowTime)) ShowTime = false;
        if (args.Length >= 7) if (!Enum.TryParse(args[7],true, out Tortoise)) Tortoise = TortoiseType.None;
        if (args.Length >= 8) if (!Enum.TryParse(args[8],true, out NamingRule)) NamingRule = NameRule.None;
    }
    
    internal static void PrintArgs()
    {
        Logger.Info($"配置表目录:{ConfigPath}");
        Logger.Info($"导出目录:{ExportPath}");
        Logger.Info($"扩展名:{Extension}");
        Logger.Info($"有序数据:{Order}");
        Logger.Info($"总表模式:{Summary}");
        Logger.Info($"注解拆分:{SplitComment}");
        Logger.Info($"文本提取:{ExtractText}");
        Logger.Info($"获取提交记录:{Tortoise}");
        Logger.Info($"显示时间流逝:{ShowTime}");
        Logger.Info($"命名规则:{NamingRule}");
        Logger.Info($"参数初始化完成");
    }

    internal static string LuaFileName(object name)
    {
        return Extension.StartsWith(".") ? $"{name}{Extension}" : $"{name}.{Extension}";
    }

    internal static string Name(string name)
    {
        return NamingRule switch
        {
            NameRule.None => name,
            NameRule.CapitalizeFirst => name.CapitalizeFirstLetter(),
            NameRule.LowercaseFirst => name.LowercaseFirstLetter(),
            NameRule.UpperCase => name.ToUpperInvariant(),
            NameRule.LowerCase => name.ToLowerInvariant(),
            NameRule.CamelCase => name.CamelCase(),
            NameRule.PascalCase => name.PascalCase(),
            _ => name
        };
    }
}