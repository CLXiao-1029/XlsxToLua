using OfficeOpenXml;
using XlsxToLua.Common;
using XlsxToLua.Data;

Console.OutputEncoding = System.Text.Encoding.UTF8;
#if NET5_0_OR_GREATER
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //指明非商业应用

try
{
    if (args.Length < 2)
    {
        Logger.ErrorAndExit("参数错误，请输入配置文件路径和导出路径");
        return;
    }
    
    MainArgs.ParseArgs(args);
    
    var cfgPath = Path.GetFullPath(MainArgs.ConfigPath);
    // 检测配置表目录是否存在
    if (!Directory.Exists(cfgPath))
    {
        Logger.ErrorAndExit("配置表目录不存在");
        return;
    }
    
    var exportPath = Path.GetFullPath(MainArgs.ExportPath);
    // 检测输出目录是否存在，不存在则创建
    if (!Directory.Exists(exportPath))
    {
        Directory.CreateDirectory(exportPath);
    }
    
    MainArgs.PrintArgs();
    
    var dateTimeAll = DateTime.Now;
    TortoiseHelper.LatestCommitRecord(MainArgs.ConfigPath);
    // 加载多语言
    MultiLanguageHelper.Load();
    // 加载配置表
    LuaExportHelper.QueryXlsxAll();
    // 导出lua表
    LuaExportHelper.XlsxToLua();
    // 导出多语言
    MultiLanguageHelper.Save();
    Logger.Info(MainArgs.ShowTime ? $"导出完成，耗时{(DateTime.Now - dateTimeAll).TotalSeconds}秒" : "导出完成");
}
catch (Exception e)
{
    Logger.Exception(e);
}