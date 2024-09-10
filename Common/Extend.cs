using System.Text;
using System.Text.RegularExpressions;

namespace XlsxToLua.Common;

public static class Extend
{
    public static void AppendLineIndent(this StringBuilder content, object data)
    {
        content.AppendLine(Values.IndentIndex() + data);
    }
    
    public static void AppendIndent(this StringBuilder content, object data)
    {
        content.Append(Values.IndentIndex() + data);
    }

    /// <summary>
    /// 移除尾随逗号，通过标志
    /// </summary>
    /// <param name="sb"></param>
    public static void RemoveTrailingComma(this StringBuilder sb)
    {
        sb.Append("^.^");
        var str1 = $",{Environment.NewLine}^.^";
        var newStr1 = $"{Environment.NewLine}";
        sb.Replace(str1, newStr1);
    }
    
    public static string CapitalizeFirstLetter(this string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // 获取第一个字符，并将其转换为大写
        var firstChar = char.ToUpper(input[0]);

        // 获取剩余的字符串
        var restOfString = input.Substring(1);

        // 将第一个字符与剩余的字符串连接起来
        return new StringBuilder().Append(firstChar).Append(restOfString).ToString();
    }
    
    public static string LowercaseFirstLetter(this string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // 获取第一个字符，并检查是否是大写字母
        if (input[0] >= 'A' && input[0] <= 'Z')
        {
            // 如果是大写字母，则转换为小写
            return char.ToLower(input[0]) + input.Substring(1);
        }
        else
        {
            // 如果第一个字符不是大写字母，则直接返回原始字符串
            return input;
        }
    }
    public static string PascalCase(this string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // 使用正则表达式来找到所有的单词
        var words = Regex.Split(input, @"[^a-zA-Z0-9]+");

        // 将所有单词的首字母转为大写，并连接起来
        for (var i = 0; i < words.Length; i++)
        {
            if (!string.IsNullOrEmpty(words[i]))
            {
                words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
            }
        }

        return string.Join("", words);
    }
    
    public static string CamelCase(this string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // 使用正则表达式来找到所有的单词
        var words = Regex.Split(input, @"[^a-zA-Z0-9]+");

        // 第一个单词首字母小写，其余单词首字母大写
        words[0] = words[0].ToLower();
        for (var i = 1; i < words.Length; i++)
        {
            if (!string.IsNullOrEmpty(words[i]))
            {
                words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
            }
        }

        return string.Join("", words);
    }
}