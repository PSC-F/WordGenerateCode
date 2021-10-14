using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;

/// <summary>
/// 名称: Word文档模板导出扩展类
/// 描述: 解析Word模板实现实体类替换导出文件流、支持表格动态追加以及表格、段落中文字内容替换。
/// </summary>
public class NPOITemplateExtensions
{
    // 是否可操作
    public static bool IsGenerate = false;

    // 类名字段
    private static string className;

    // 循环类名
    private static string loopName;

    // 动态类名
    private static string dynamicClassName;

    // 维护数据字典
    public static Dictionary<string, dynamic> DataDictionary = new Dictionary<string, dynamic>();

    // 维护动态类型数据类名
    public static List<string> DynamicDataClassName = new List<string>();

    // 动态类型
    public static Dictionary<string, Dictionary<int, int>> TableParaInfo =
        new Dictionary<string, Dictionary<int, int>>();

    /// <summary>
    /// 添加待导出文档的实体
    /// </summary>
    /// <param name="o">XXEntity</param>
    public static void AddEntity(dynamic o)
    {
        Type t = o.GetType();
        if (!string.IsNullOrEmpty(t.FullName))
        {
            if (DataDictionary.ContainsKey(t.FullName))
            {
                DataDictionary.Remove(t.FullName);
            }

            DataDictionary.Add(t.FullName, o);
        }
    }

    /// <summary>
    /// 添加待导出文档的动态类型数据
    /// </summary>
    /// <param name="key">类名与文档标签对应</param>
    /// <param name="o">动态数据List<XXXEntity></param>
    public static void AddDynamicData(string key, dynamic o)
    {
        if (!string.IsNullOrEmpty(key))
        {
            if (DataDictionary.ContainsKey(key))
            {
                DataDictionary.Remove(key);
            }

            DataDictionary.Add(key, o);
        }
    }

    /// <summary>
    /// 输出模板docx文档(使用字典) /// </summary>
    /// <param name="tempFilePath">docx文件路径</param>
    /// <param name="outPath">输出文件路径</param>
    /// <param name="data">字典数据源</param>
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

            // 写入流
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
    /// 是否可生成类
    /// </summary>
    /// <param name="out_className">类名</param>
    /// <param name="dynamic_className">动态表类型类名</param>
    /// <param name="loop_name">循环类型类名</param>
    public static bool IsGenerateEntityClass(
        XWPFParagraph para,
        out string out_className,
        out string dynamic_className,
        out string loop_name
    )
    {
        // 匹配是有开始标记、有则创建类文件、并遍历模板
        // 如果匹配到结束标记、break;
        // 发现标记 实例化类、替换并写入对象名;
        Regex s = new Regex(@"<[s=]+.*?>"); //开始标记
        Regex d = new Regex(@"<[d=]+.*?>"); //动态标记
        Regex loop = new Regex(@"<[loop=]+.*?>"); //动态标记
        var className = "";
        var dynamicClassName = "";
        var LoopName = "";
        var IsGenerate = s.IsMatch(para.Text);
        var IsDynamic = d.IsMatch(para.Text);
        var IsLoop = loop.IsMatch(para.Text);
        if (IsGenerate)
        {
            var str = s.Match(para.Text).Value;
            className = str.Substring(3, str.Length - 4);
            //删除 <s=>标签
            para.ReplaceText(str, "");
        }

        if (IsDynamic)
        {
            var str = d.Match(para.Text).Value;
            dynamicClassName = str.Substring(3, str.Length - 4);
            //删除 <s=>标签
            para.ReplaceText(str, "");
        }

        if (IsLoop)
        {
            var str = loop.Match(para.Text).Value;
            LoopName = str.Substring(6, str.Length - 7);
            para.ReplaceText(str, "");
        }

        loop_name = LoopName;
        out_className = className;
        dynamic_className = dynamicClassName;
        return IsGenerate;
    }

    /// <summary>
    /// 使用普通对象导出Word文档
    /// </summary>
    /// <param name="inStream"></param>
    /// <param name="data"></param>
    /// <returns></returns>
    public static Stream ExportObjet(Stream inStream, object data)
    {
        XWPFDocument doc = new XWPFDocument(inStream);
        // 遍历段落 
        foreach (var para in doc.Paragraphs)
        {
            ReplaceKeyObjet(para, data);
        }

        // 遍历表格 
        foreach (var table in doc.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        ReplaceKeyObjet(para, data);
                    }
                }
            }
        }

        // 输出流
        MemoryStream outputStream = new MemoryStream();
        doc.Write(outputStream);
        MemoryStream memoryStream = new MemoryStream(outputStream.ToArray());
        return memoryStream;
    }

    /// <summary>
    /// 重载文档
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="outputStream"></param>
    /// <returns></returns>
    public static XWPFDocument reloadDoc(XWPFDocument doc, MemoryStream outputStream)
    {
        doc.Write(outputStream);
        MemoryStream memoryStream = new MemoryStream(outputStream.ToArray());
        return new XWPFDocument(memoryStream);
    }

    /// <summary>
    /// 使用动态数据导出word文档
    /// </summary>
    /// <param name="inStream"></param>
    /// <returns></returns>
    public static Stream ExportWithDynamicList(Stream inStream)
    {
        XWPFDocument doc = new XWPFDocument(inStream);
        MemoryStream outputStream = new MemoryStream();
        foreach (var para in doc.Paragraphs)
        {
            if (!IsGenerate) // 开锁
            {
                // 是否可生成类
                if (IsGenerateEntityClass(
                    para,
                    out string className,
                    out string dynamicClassName,
                    out string loopName)
                )
                {
                    if (!string.IsNullOrEmpty(className))
                    {
                        NPOITemplateExtensions.className = className;
                        NPOITemplateExtensions.dynamicClassName = dynamicClassName;
                    }
                }

                // 渲染word模板静态数据
                if (!string.IsNullOrEmpty(NPOITemplateExtensions.className))
                {
                    if (DataDictionary.TryGetValue(className, out dynamic entity))
                    {
                        foreach (var par in doc.Paragraphs)
                        {
                            ReplaceKeyObjet(par, entity);
                        }
                    }
                }

                // 渲染word模板静态数据
                if (!string.IsNullOrEmpty(loopName))
                {
                    if (DataDictionary.TryGetValue(loopName, out dynamic entity))
                    {
                        for (int i = 0; i < entity.Count; i++)
                        {
                            StringBuilder rowText = new StringBuilder();
                            var run = para.CreateRun();
                            foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(entity[i]))
                            {
                                try
                                {
                                    // 拼接新的一行文本数据
                                    var text = entity[i]
                                        .GetType()
                                        .GetProperty(prop.Name)
                                        .GetValue(entity[i], null) + "";
                                    rowText.Append(text);
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }

                            // 每段添加Tab缩进
                            if (!string.IsNullOrEmpty(rowText.ToString()))
                            {
                                run.SetText(rowText.ToString());
                                run.AddCarriageReturn();
                                if (i >= 0)
                                {
                                    run.AddTab();
                                }
                            }
                        }
                    }
                }

                if (IsHasEndKey(para))
                {
                    IsGenerate = false; // 关锁
                    NPOITemplateExtensions.className = "";
                }
            }
        }

        // 重置状态
        IsGenerate = false;
        // 清空
        className = "";
        // 遍历表格 
        foreach (var table in doc.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        if (!IsGenerate) // 开锁
                        {
                            // 是否可生成类
                            if (IsGenerateEntityClass(
                                para,
                                out string className,
                                out string dynamicClassName,
                                out string loopName))
                            {
                                if (!string.IsNullOrEmpty(className))
                                {
                                    NPOITemplateExtensions.className = className;
                                }
                            }

                            {
                                if (!string.IsNullOrEmpty(dynamicClassName))
                                {
                                    Console.WriteLine(dynamicClassName);
                                    NPOITemplateExtensions.dynamicClassName = dynamicClassName;
                                    if (DynamicDataClassName.Contains(dynamicClassName))
                                    {
                                        DynamicDataClassName.Remove(dynamicClassName);
                                    }

                                    DynamicDataClassName.Add(dynamicClassName);
                                }
                            }

                            // 如果存在实体类生成字段
                            if (!string.IsNullOrEmpty(NPOITemplateExtensions.className))
                            {
                                if (DataDictionary.TryGetValue(NPOITemplateExtensions.className, out dynamic entity))
                                {
                                    ReplaceKeyObjet(para, entity);
                                }
                            }

                            if (IsHasEndKey(para))
                            {
                                IsGenerate = false; // 关锁
                                NPOITemplateExtensions.className = "";
                            }
                        }
                    }
                }
            }
        }

        // word_动态数据表格
        foreach (var className in DynamicDataClassName)
        {
            if (DataDictionary.TryGetValue(className, out dynamic entity))
            {
                try
                {
                    // 获取表格索引
                    string code = entity[0].GetType().GetProperty("TableIndex").GetValue(entity[0], null).ToString();
                    // 追加行
                    if (int.TryParse(code, out int codeNum))
                    {
                        XWPFTable dynamicTable = doc.Tables[codeNum];
                        // 遍历List
                        for (int i = 0; i < entity.Count - 1; i++)
                        {
                            // 获取最后一行
                            var endRow = dynamicTable.Rows[dynamicTable.Rows.Count - 1];
                            // 表尾创建新行
                            var xwpfTableCells = endRow.GetTableCells();
                            var r = dynamicTable.CreateRow();
                            r.GetCTRow().AddNewTrPr().AddNewTrHeight().val = (ulong) 426;
                            foreach (var xwpfTableCell in r.GetTableCells())
                            {   
                                // 设置垂直居中
                                xwpfTableCell.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                                // 设置水平居中
                                var ctTc = xwpfTableCell.GetCTTc();
                                ctTc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.left;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        // 赋值
        foreach (var className in DynamicDataClassName)
        {
            if (DataDictionary.TryGetValue(className, out dynamic entity))
            {
                try
                {
                    // 获取表格索引
                    string code = entity[0].GetType().GetProperty("TableIndex").GetValue(entity[0], null).ToString();
                    // 追加行
                    if (int.TryParse(code, out int codeNum))
                    {
                        XWPFTable dynamicTable = doc.Tables[codeNum];
                        // 遍历集合
                        for (int rowIndex = 0; rowIndex < entity.Count; rowIndex++)
                        {
                            var colIndex = 0;
                            // 遍历对象属性
                            int FiledCount = 0;
                            foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(entity[rowIndex]))
                            {
                                FiledCount++;
                            }

                            foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(entity[rowIndex]))
                            {
                                if (prop.Name.Equals("TableIndex"))
                                {
                                    break;
                                }
                                try
                                {
                                    // Console.WriteLine(rowIndex + "行");
                                    // Console.WriteLine(colIndex + "列");
                                    // Console.WriteLine(dynamicTable.Rows.Count+"数量");
                                    // Console.WriteLine(entity.Count+"列表数量");
                                    var startIndex = 1 + rowIndex;
                                    var p = dynamicTable.Rows[startIndex]
                                        .GetCell(colIndex)
                                        .Paragraphs[0];
                                    Console.WriteLine(p.Text + "单元格对象");
                                    if (!string.IsNullOrEmpty(p.Text))
                                    {
                                        p.ReplaceText(
                                            p.Text,
                                            entity[rowIndex]
                                                .GetType()
                                                .GetProperty(prop.Name)
                                                .GetValue(entity[rowIndex], null).ToString());
                                    }
                                    else
                                    {
                                        // 为动态创建的空白行追加数据
                                        var cellsList = dynamicTable.Rows[startIndex].GetTableCells();
                                        Console.WriteLine(cellsList.Count);
                                        for (int i = 0; i < cellsList.Count; i++)
                                        {
                                            // 匹配列与属性
                                            if (i == colIndex)
                                            {
                                                Console.WriteLine(entity[rowIndex]
                                                    .GetType()
                                                    .GetProperty(prop.Name)
                                                    .GetValue(entity[rowIndex], null).ToString() + "列" + i);
                                                cellsList[i].SetText(
                                                    entity[rowIndex]
                                                        .GetType()
                                                        .GetProperty(prop.Name)
                                                        .GetValue(entity[rowIndex], null).ToString());
                                            }
                                        }
                                    }

                                    // 去掉属性中的TableIndex占位
                                    if (colIndex < FiledCount - 2)
                                    {
                                        colIndex++;
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        // 重载流
        doc = reloadDoc(doc, outputStream);
        // 创建新的输出流
        outputStream = new MemoryStream();
        // 表格纵向合并  获取行、获取列
        // 找到动态表 传入参数
        // foreach (var className in DynamicDataClassName)
        // {
        //     if (DataDictionary.TryGetValue(className, out dynamic entity))
        //     {
        //         try
        //         {
        //             // 获取表格索引
        //             string code = entity[0].GetType().GetProperty("TableIndex").GetValue(entity[0], null).ToString();
        //             if (int.TryParse(code, out int codeNum))
        //             {}}
        // 按相同内容纵向合并单元格(暂时指定0列)
        // 取表头
        var tableTotalCount = doc.Tables.Count;
        for (int tableIndex = 0; tableIndex < tableTotalCount; tableIndex++)
        {
            var row = doc.Tables[tableIndex].GetRow(0);
            // 获取列数
            int colNums = 0;
            foreach (var xwpfTableCell in row.GetTableCells())
            {
                colNums++;
            }

            // 获取行数
            var rowNums = doc.Tables[tableIndex].NumberOfRows;
            int fromNum = -1;
            int toNum = 0;
            for (int rowIndex = 1; rowIndex <= rowNums - 2; rowIndex++)
            {
                var currText = doc.Tables[tableIndex]
                    .GetRow(rowIndex)
                    .GetCell(0)
                    .GetText();
                var nextText = doc.Tables[tableIndex]
                    .GetRow(rowIndex + 1)
                    .GetCell(0)
                    .GetText();
                if (currText.Equals(nextText))
                {
                    if (fromNum == -1)
                    {
                        fromNum = rowIndex; // 起始索引
                    }

                    ++toNum;
                }
                // 计算第一列
            }

            // 纵向合并单元格
            if (fromNum != -1)
            {
                mergeCellVert(doc.Tables[tableIndex], 0, fromNum, toNum + 1);
            }
        }


        // 输出流
        // MemoryStream outputStream = new MemoryStream();
        doc.Write(outputStream);
        MemoryStream memoryStream = new MemoryStream(outputStream.ToArray());
        return memoryStream;
    }

    /// <summary>
    /// 流转文件
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="fileName"></param>
    public static void StreamToFile(Stream stream, string fileName)
    {
        // 把 Stream 转换成 byte[]   
        byte[] bytes = new byte[stream.Length];
        stream.Read(bytes, 0, bytes.Length);
        // 设置当前流的位置为流的开始   
        stream.Seek(0, SeekOrigin.Begin);

        // 把 byte[] 写入文件   
        FileStream fs = new FileStream(fileName, FileMode.Create);
        BinaryWriter bw = new BinaryWriter(fs);
        bw.Write(bytes);
        bw.Close();
        fs.Close();
    }

    /// <summary>
    /// 是否包含结束标签
    /// </summary>
    /// <param name="para"></param>
    /// <returns></returns>
    public static bool IsHasEndKey(XWPFParagraph para)
    {
        Regex s = new Regex(@"<[e]+.*?>"); //开始标记
        var str = s.Match(para.Text).Value;
        //删除结束<e>标记
        if (s.IsMatch(para.Text))
        {
            para.ReplaceText(str, "");
        }

        return s.IsMatch(para.Text);
    }

    /// <summary>
    /// 根据对象替换模板内容
    /// 
    /// </summary>
    /// <param name="para"></param>
    /// <param name="model"></param>
    private static void ReplaceKeyObjet(XWPFParagraph para, dynamic model)
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
                string key = $"${p.Name}$";
                if (text.Contains(key))
                {
                    try
                    {
                        text = text.Replace(key, p.GetValue(model, null).ToString());
                    }
                    catch (Exception ex)
                    {
                        // 可能有空指针异常
                        text = text.Replace(key, "");
                    }
                }
                else if (text.Contains($@"#{p.Name}#"))
                {
                    // var gfs = new FileStream($@"D:\\UploadFiles\\{p.Name}.jpg", FileMode.Open, FileAccess.Read);
                    text = text.Replace($@"#{p.Name}#", p.GetValue(model, null).ToString());
                    try
                    {
                        // bug here Can not normal display  
                        Stream stream = model.GetType().GetProperty(p.Name).GetValue(model, null);
                        run.AddPicture(stream, (int) NPOI.XWPF.UserModel.PictureType.PNG, $@"{p.Name}.PNG", 5300000,
                            2500000);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }

            run.SetText(text, 0);
        }
    }

    /// <summary>
    /// 纵向合并Table单元格
    /// </summary>
    /// <param name="table">目标表</param>
    /// <param name="row">指定列索引</param>
    /// <param name="fromCell">指定起始行</param>
    /// <param name="toCell">指定结束行</param>
    public static void mergeCellVert(XWPFTable table, int colIndex, int fromCell, int toCell)
    {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++)
        {
            XWPFTableCell cell = table.GetRow(cellIndex).GetCell(colIndex);
            if (cellIndex == fromCell)
            {
                Console.WriteLine("Restart: " + cell.GetText());
                cell.GetCTTc().AddNewTcPr().AddNewVMerge().val = ST_Merge.restart;
            }
            else
            {
                Console.WriteLine("continue: " + cell.GetText());

                cell.GetCTTc().AddNewTcPr().AddNewVMerge().val = ST_Merge.@continue;
            }
        }
    }
}