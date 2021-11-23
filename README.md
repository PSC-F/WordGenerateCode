# WordGenerateCode🚀
* 本项目基于NPOI实现Word文档模板与实体类映射,通过在Word(模板)中添加文本替换标签<s=className>$filedd$ <e>、或 #图片# 标签, 重新对word模板文档内容进行结构化、结合业务生成包含文本、动态自适应表格、图片的Word文档、高效生成复杂Word文档☺。
## example
 * 1.制作Word 模板
 * 2.结构化模板、声明模板语法 
 <s=ClassName>  // 开始标记
  $filed$  // 动态输出内容 (支持整段落、或单句)
 <e>   // 结束标记
  ---
  组合示例:
 <s=Company>$Pretty$</e>
  ---
  *3.调用实体类生成工具 即可生成相应实体类代码
  output->
  Company.cs
  ```
  -- public dynamic Pretty { get; set; }
  ```
  ---
  *tips：可搭配OSS云存储使用、模板放在云端
* OSS ☁ ->pull template🎫 
*  -> Replace🚗
*  -> generate newWord 🎫 
*  <- push file 🎫 
*  -> dowload url 🎯
* 4.代码调用、结合业务调用AddEntity、AddDynamicData填充数据、 调用生成ExportWithDynamicList方法即可生成复杂文档。
 ``` C#
  NPOITemplateExtensions.AddEntity(BridgeConcent); // 添加实体
  NPOITemplateExtensions.AddEntity(BridgeMember);
  NPOITemplateExtensions.AddEntity(BridgeOverview); 
  NPOITemplateExtensions.AddDynamicData("SurfaceDiseaseList",SurfaceDiseaseList); // 添加动态数据实体
  NPOITemplateExtensions.ExportWithDynamicList(templateStreamFromOss) // 动态导出
``` 
