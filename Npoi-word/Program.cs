using System;
using System.IO;
using NPOI.XWPF.UserModel;


// OSS templateName
// param : {templateName,Datas:[]},
// 


namespace Npoi_word
{
    class Program
    {
        static void Main(string[] args)
        {
            NPOITemplateExtensions.ExportObjet(
                "d://UploadFiles//广东省农村桥梁评定报告.docx",
                "d://UploadFiles//result.docx",
                new
                {
                    pic = "",
                    数据1="XXXXXXXXX"
                }
                // 模板语法 #图片名#   ||   $文本内容$                
            );
        }
    }
}