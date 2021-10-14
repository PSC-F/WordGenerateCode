using System;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.XWPF.UserModel;

namespace Npoi_word.DocExtions
{
    public class Tools
    {
        /// <summary>
        /// 生成类文件
        /// </summary>
        /// <param name="className">类名</param>
        /// <param name="content">写入内容</param>
        public static void Write(String className, String content)
        {
            File.AppendAllText($@"{className}.cs", "\r\n" + content);
        }

        /// <summary>
        /// 是否可生成类
        /// </summary>
        /// <param name="para"></param>
        public static bool IsGenerateEntityClass(XWPFParagraph para, out string out_className)
        {
            // 匹配是有开始标记、有则创建类文件、并遍历模板
            // 如果匹配到结束标记、break;
            // 发现标记 实例化类、替换并写入对象名;
            Regex s = new Regex(@"<[s=]+.*?>"); //开始标记
            var className = "";
            var IsGenerate = s.IsMatch(para.Text);
            if (IsGenerate)
            {
                var str = s.Match(para.Text).Value;
                className = str.Substring(3, str.Length - 4);
                Tools.Write(
                         className,
                $@"
/// <summary>
/// desc: {className}
/// time: {DateTime.Now.ToString()}
/// remark: Word文档代码生成器生成、DTO对象
/// au: zpf  
/// </summary>
public class {className}{{
                    ");
            }

            out_className = className;
            return IsGenerate;
        }

        public static void GenerateEntity(string className, XWPFParagraph para)
        {
            Regex text = new Regex(@"[$]([\s\S]*?)[$]"); //文本
            Regex img = new Regex(@"[#]([\s\S]*?)[#]"); //图片
            Regex e = new Regex(@"<[e]+.*?>"); //结束标记

            if (text.IsMatch(para.Text))
            {
                // 生成图片Fileds字段
                foreach (Match match in text.Matches(para.Text))
                {
                    var filed = match.Value;
                    Tools.Write(className,
                        $@"
   public dynamic {filed.Substring(1, filed.Length - 2)} {{ get; set; }}"
                        );
                }
            }
            else if (img.IsMatch(para.Text))
            {
                // 生成图片Fileds字段
                foreach (Match match in img.Matches(para.Text))
                {
                    var filed = match.Value;
                    Tools.Write(className,
                        $@"
   public dynamic {filed.Substring(1, filed.Length - 2)} {{ get; set; }}"
                    );
                }
            }
            if (e.IsMatch(para.Text))
            {   // 匹配到结束标记 表示结束
                Tools.Write(className,
                    $@"
}}");
            }
        }

        public static bool IsHasStartKey(XWPFParagraph para)
        {  Regex s = new Regex(@"<[s=]+.*?>"); //开始标记
           return s.IsMatch(para.Text);
        }

        public static bool IsHasEndKey(XWPFParagraph para)
        {
            Regex s = new Regex(@"<[e]+.*?>"); //开始标记
            return s.IsMatch(para.Text);
        }
    }
}