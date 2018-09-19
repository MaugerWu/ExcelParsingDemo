# ExcelParsingDemo

## POI 3.9 解析 Excel（.xsl & .xslx）

POI简介：Jakarta POI 是一套用于访问微软格式文档的Java API。Jakarta POI有很多组件组成，其中有用于操作Excel格式文件的HSSF和用于操作Word的HWPF，在各种组件中目前只有用于操作Excel的HSSF相对成熟。

官方主页http://poi.apache.org/index.html，

API文档http://poi.apache.org/apidocs/index.html

HSSF（用于操作Excel的组件）提供给用户使用的对象在rg.apache.poi.hssf.usermodel包中,主要部分包括Excel对象，样式和格式，有以下几种常用的对象：

常用组件：

HSSFWorkbook     excel的文档对象

HSSFSheet            excel的表单

HSSFRow               excel的行

HSSFCell                excel的格子单元

HSSFFont               excel字体

样式：

HSSFCellStyle         cell样式

## 所需 jar 包

1. poi-3.9.jar
2. poi-ooxml-3.8.jar
3. poi-ooxml-schemas-3.8.jar
4. xmlbeans-2.6.0.jar
5. commons-collections4-4.0.jar
6. dom4j-1.6.1.jar
