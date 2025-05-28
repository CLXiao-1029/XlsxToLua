using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using XlsxToLua.Common;

namespace XlsxToLua.Data;

internal static class TortoiseHelper
{
    public static string? CommitLog; 
    internal static void LatestCommitRecord(string workPath)
    {
        switch (MainArgs.Tortoise)
        {
            case TortoiseType.Git:
                GitCommand("log -1 --pretty=%h",workPath,out CommitLog);
                break;
            case TortoiseType.Svn:
            {
                SvnCommand("info",workPath,out var info);
                const string pattern = @"Revision:\s*(\d+)";
                var match = Regex.Match(info, pattern);
                if (match.Success)
                    CommitLog = match.Groups[1].Value;

                break;
            }
            default:
                CommitLog = null;
                break;
        }
        
        Logger.Warning($"最新提交记录: {CommitLog}");
    }

    private static void GitCommand(string command, string workingDirectory, out string line)
    {
        var fileName = "git";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            fileName = "git.exe";
        }
        var p = new Process();
        p.StartInfo.FileName = fileName;
        p.StartInfo.Arguments = command;
        p.StartInfo.WorkingDirectory = workingDirectory;
        p.StartInfo.CreateNoWindow = true;
        p.StartInfo.UseShellExecute = false;
        p.StartInfo.RedirectStandardOutput = true;
        p.StartInfo.RedirectStandardInput = true;
        p.StartInfo.RedirectStandardError = true;
        p.StartInfo.StandardOutputEncoding = Encoding.UTF8;
        p.Start();
        line = p.StandardOutput.ReadToEnd();
        p.WaitForExit();
        p.Close();
        p.Dispose();
    }
    
    private static void SvnCommand(string command, string workingDirectory, out string line)
    {
        var fileName = "svn";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            fileName = "svn.exe";
        }
        using var process = new Process();
        process.StartInfo.FileName = fileName;
        process.StartInfo.Arguments = command;
        process.StartInfo.WorkingDirectory = workingDirectory;
        process.StartInfo.CreateNoWindow = true;
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardInput = true;
        process.StartInfo.RedirectStandardError = true;
        process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
        process.Start();
        line = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        process.Close();
        process.Dispose();
    }
}