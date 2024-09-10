namespace XlsxToLua.Common;

internal class FileUtils
{
    /// <summary>
    /// 获取指定目录下的所有文件（包含子目录）
    /// </summary>
    /// <param name="path"></param>
    /// <param name="files"></param>
    internal static void GetDirectoryFiles(string path, ref List<FileInfo> files)
    {
        files.AddRange(GetFiles(path));
        var directoryInfo = new DirectoryInfo(path);
        foreach (var directory in directoryInfo.GetDirectories())
        {
            GetDirectoryFiles(directory.FullName, ref files);
        }
    }

    /// <summary>
    /// 获取顶部目录下的所有文件（包含子目录），并以顶部目录名作为Key记录下来
    /// </summary>
    /// <param name="path"></param>
    /// <param name="filesDic"></param>
    internal static void GetTopDirectoryFiles(string path, ref Dictionary<string, List<FileInfo>> filesDic)
    {
        var directoryInfo = new DirectoryInfo(path);
        foreach (var directory in directoryInfo.GetDirectories())
        {
            var folder = directory.Name;
            var files = new List<FileInfo>();
            GetDirectoryFiles(Path.Combine(path, folder), ref files);
            filesDic.Add(folder, files);
        }
    }

    internal static IEnumerable<FileInfo> GetFiles(string path, string folder = "")
    {
        if (folder != "")
            path = Path.Combine(path, folder);

        var directoryInfo = new DirectoryInfo(path);
        var fileInfos = directoryInfo.GetFiles().OrderBy(file => file.Name);
        return fileInfos.ToArray();
    }
    
    private static void SafeCreateDirectory(string? path)
    {
        if (path == null)
            return;
        
        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }
    }
    
    #region 文件读写
    
    internal static bool SafeSave(string fileName, string content)
    {
        try
        {
            if (MainArgs.ExportPath != string.Empty)
                fileName = Path.Combine(MainArgs.ExportPath, fileName);
            
            SafeCreateDirectory(Path.GetDirectoryName(fileName));

            File.WriteAllText(fileName, content);
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }
        return false;
    }

    #endregion
}