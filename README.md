# WordGenerateCode
* 本项目基于NPOI实现Word文档模板->实体类的代码生成器,通过在word模板中添加<s=className>$filedd$<e>语法格式、重新对word文档内容进行结构化、搭配模板替换工具、实现word文档导出功能。
## example
  * <s=Company>$Pretty$> <e>
  
  output->
  Company.cs
  -- public dynamic Pretty { get; set; }
  
