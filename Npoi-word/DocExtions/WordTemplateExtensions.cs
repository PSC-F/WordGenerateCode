using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using NPOI.XWPF.UserModel;
using Npoi_word.DocExtions;

/// <summary>
/// NPOI模板导出转换扩展
/// </summary>
public class NPOITemplateExtensions
{
    /// <summary>
    /// 输出模板docx文档(使用字典) /// </summary>
    /// <param name="tempFilePath">docx文件路径</param>
    /// <param name="outPath">输出文件路径</param>
    /// <param name="data">字典数据源</param>
    public static bool IsGenerate = false;
    // 待生成类名
    public static string className = "";
    public static void Export(string tempFilePath, string outPath, Dictionary<string, string> data)
    {
        using (FileStream stream = File.OpenRead(tempFilePath))
        {
            XWPFDocument doc = new XWPFDocument(stream);
            //遍历段落 
            foreach (var para in doc.Paragraphs)
            {
                ReplaceKey(para, data);
            }

            //遍历表格 
            foreach (var table in doc.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        foreach (var para in cell.Paragraphs)
                        {
                            ReplaceKey(para, data);
                        }
                    }
                }
            }

            // 写入
            FileStream outFile = new FileStream(outPath, FileMode.Create);
            doc.Write(outFile);
            outFile.Close();
        }
    }

    /// <summary>
    /// 替换模板key
    /// </summary>
    /// <param name="para">段落</param>
    /// <param name="data">字典</param>
    private static void ReplaceKey(XWPFParagraph para, Dictionary<string, string> data)
    {
        string text = "";

        foreach (var run in para.Runs)
        {
            text = run.ToString();
            foreach (var key in data.Keys)
            {
                // $$模板中数据占位符为$KEY$
                if (text.Contains($"${key}$"))
                {
                    text = text.Replace($"${key}$", data[key]);
                }
            }

            run.SetText(text, 0);
        }
    }

    /// <summary>
    /// 输出模板docx文档(使用反射) /// </summary>
    /// <param name="tempFilePath">docx文件路径</param>
    /// <param name="outPath">输出文件路径</param>
    /// <param name="data">对象数据源</param>
    public static void ExportObjet(string tempFilePath, string outPath, object data)
    {
        using (FileStream stream = File.OpenRead(tempFilePath))
        {
            XWPFDocument doc = new XWPFDocument(stream);
            // 遍历段落 
            foreach (var para in doc.Paragraphs)
            {   // 是否有开始标记
                if (!IsGenerate) // 开锁
                {
                    if (!IsGenerate) // 开锁
                    {
                        // 是否可生成类
                        if (Tools.IsGenerateEntityClass(para,out string className))
                        {
                            if (!string.IsNullOrEmpty(className))
                            {
                                NPOITemplateExtensions.className = className;
                            }
                        }
                        // 如果存在实体类生成字段
                        if (!string.IsNullOrEmpty(NPOITemplateExtensions.className))
                        {
                            Tools.GenerateEntity(NPOITemplateExtensions.className,para);
                        }
                        if (Tools.IsHasEndKey(para))
                        {
                            IsGenerate = false; //关锁
                            NPOITemplateExtensions.className = "";
                        }
                    }
                }
            }

            IsGenerate = false;
            //遍历表格 
            foreach (var table in doc.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        foreach (var para in cell.Paragraphs)
                        {
                            if (!IsGenerate) // 开锁
                            {   // 是否可生成类
                                if (Tools.IsGenerateEntityClass(para,out string className))
                                {
                                    if (!string.IsNullOrEmpty(className))
                                    {
                                        NPOITemplateExtensions.className = className;
                                    }
                                }
                                // 如果存在实体类生成字段
                                if (!string.IsNullOrEmpty(NPOITemplateExtensions.className))
                                {
                                    Tools.GenerateEntity(NPOITemplateExtensions.className,para);
                                }
                                if (Tools.IsHasEndKey(para))
                                {
                                    IsGenerate = false; // 关锁
                                    NPOITemplateExtensions.className = "";
                                }
                            }
                        }
                    }
                }
            }

            //写文件
            FileStream outFile = new FileStream(outPath, FileMode.Create);
            doc.Write(outFile);
            outFile.Close();
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="para"></param>
    /// <param name="model"></param>
    private static void ReplaceKeyObjet(XWPFParagraph para, object model)
    {
        string text = "";
        Type t = model.GetType();
        PropertyInfo[] pi = t.GetProperties();
        foreach (var run in para.Runs)
        {
            text = run.ToString();
            foreach (PropertyInfo p in pi)
            {
                //$$模板中数据占位符为$KEY$
                // string startKey = $"<s=/>$"; //开始标记 =定义描述
                // string endKey = "</>"; // 结束标记
                string textkey = $"${p.Name}$"; // 文本标记
                string imgKey = $@"#{p.Name}#"; // 图片标记
                Console.WriteLine(textkey);
                // 如果包含开始标记
                if (text.Contains(textkey))
                {
                    try
                    {
                        text = text.Replace(textkey, p.GetValue(model, null).ToString());
                    }
                    catch (Exception ex)
                    {
                        // 可能有空指针异常
                        text = text.Replace(textkey, "");
                    }
                }
                else if (text.Contains(imgKey))
                {
                    var gfs = new FileStream($@"D:\\UploadFiles\\{p.Name}.jpg", FileMode.Open, FileAccess.Read);
                    text = text.Replace($@"#{p.Name}#", p.GetValue(model, null).ToString());
                    run.AddPicture(gfs, (int) NPOI.XWPF.UserModel.PictureType.JPEG, $@"{p.Name}.jpg", 5300000, 2500000);
                    gfs.Close();
                }
                else if (text.Contains(imgKey))
                {
                    var gfs = new FileStream($@"D:\\UploadFiles\\{p.Name}.jpg", FileMode.Open, FileAccess.Read);
                    text = text.Replace($@"#{p.Name}#", p.GetValue(model, null).ToString());
                    run.AddPicture(gfs, (int) NPOI.XWPF.UserModel.PictureType.JPEG, $@"{p.Name}.jpg", 5300000, 2500000);
                    gfs.Close();
                }
            }


            run.SetText(text, 0);
        }
    }
}