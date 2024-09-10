using System.Text;
namespace XlsxToLua.Common;

internal class Logger
{
    private static readonly StringBuilder LogContent = new();
    
    internal static void Info(object message)
    {
        Console.ForegroundColor = ConsoleColor.White;
        Log($"LogInfo:{message}");
    }

    internal static void Error(object message)
    {
        Console.ForegroundColor = ConsoleColor.DarkRed;
        Log($"LogError:{message}");
        Console.ForegroundColor = ConsoleColor.White;
    }

    internal static void Warning(object message)
    {
        Console.ForegroundColor = ConsoleColor.DarkYellow;
        Log($"LogWarning:{message}");
        Console.ForegroundColor = ConsoleColor.White;
    }
    
    internal static void Exception(Exception message)
    {
        Console.ForegroundColor = ConsoleColor.DarkRed;
        Log($"LogException:{message}");
        Log($"LogException:程序被迫退出，请修正错误后重试");
        Console.ForegroundColor = ConsoleColor.White;
        Environment.Exit(0);
    }

    /// <summary>
    /// 输出错误信息并在用户按任意键后退出
    /// </summary>
    internal static void ErrorAndExit(object message)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Log($"LogError:{message}");
        Log($"LogError:程序被迫退出，请修正错误后重试");
        Console.ForegroundColor = ConsoleColor.White;
        Console.ReadKey();
        Environment.Exit(0);
    }
    
    public override string ToString()
    {
        return LogContent.ToString();
    }

    private static void Log(object message)
    {
        var msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";
        if (MainArgs.Summary)
            msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 【总表模式】 {message}";
        Console.WriteLine(msg);
        LogContent.AppendLine(msg);
    }

    internal Logger(object message)
    {
        Info(message);
    }
}