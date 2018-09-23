# ExcelParsingDemo

## POI 3.9 解析 Excel（.xsl & .xslx）

POI 简介：Jakarta POI 是一套用于访问微软格式文档的 Java API。Jakarta POI 有很多组件组成，其中有用于操作 Excel 格式文件的 HSSF 和用于操作 Word 的 HWPF，在各种组件中目前只有用于操作 Excel 的 HSSF 相对成熟。

官方主页: http://poi.apache.org/index.html，

API 文档: http://poi.apache.org/apidocs/index.html

HSSF（用于操作 Excel 的组件）提供给用户使用的对象在`rg.apache.poi.hssf.usermodel`包中,主要部分包括 Excel 对象、样式和格式，有以下几种常用的对象：

常用组件：

HSSFWorkbook          excel 的文档对象

HSSFSheet             excel 的表单

HSSFRow               excel 的行

HSSFCell              excel 的格子单元

HSSFFont              excel 字体

## 所需 jar 包

1. poi-3.9.jar
2. poi-ooxml-3.8.jar
3. poi-ooxml-schemas-3.8.jar
4. xmlbeans-2.6.0.jar
5. commons-collections4-4.0.jar
6. dom4j-1.6.1.jar
