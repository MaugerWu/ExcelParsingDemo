# ExcelParsingDemo

## POI 3.9 解析 Excel（.xsl & .xslx）

POI 简介：[POI](http://poi.apache.org/) 是一套用于访问微软格式文档的 Java API。Jakarta POI 有很多组件组成，其中有用于操作 Excel 格式文件的 HSSF 和用于操作 Word 的 HWPF，在各种组件中目前只有用于操作 Excel 的 HSSF 相对成熟。

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

## HSSFWorkbook、XSSFWorkbook 与 SXSSFWorkbook 的区别

#### HSSFWorkbook

HSSFWorkbook 针对的是 Excel 2003 的版本，扩展名为`.xls`，导出的行数 至多为 65535 行，发现只要是 Excel 文件大于 2M 左右，便会出现 OOM（Out Of Memory）；

#### XSSFWorkbook

这种形式的出现 是由于 第一种HSSFWorkbook 的局限性而产生的，因为其所导出的行数比较少，所以 XSSFWookbook应运而生 其 对应的是EXCEL2007+(1048576行，16384列)扩展名
 .xlsx，最多可以导出 104 万行，不过 这样 就伴随着一个问题---OOM 内存溢出，原因是 你所 创建的 workbook sheet row cell 等 此时是存在 内存的 并没有 持久化，那么 随着 数据量增大  内存的需求量也就增大，那么很大可能就是要 OOM了，那么 怎么解决呢？

#### SXSSFWorkbook　　poi.jar 3.8+
第二种遇到的问题该如何解决呢？ 因为数据量过大 导致内存吃不消 那么 可以 让内存 到量持久化 吗？ 
答案是 肯定的，
此种的情况就是设置最大内存条数，比如设置最大内存量为 5000 rows -- new SXSSFWookbook(5000); 此时当行数达到 5000 时，把内存持久化写到文件中，以此逐步写入，避免 OOM，那么这样就完美解决了大数据下导出的问题；

经过查询得知，原来 POI 读取 Excel 的原理如下：`org.apache.poi.xssf.usermodel.XSSFWorkbook.XSSFWorkbook(InputStream is) throws IOException` 采用`usermodel`，这种方式是以`dom`方式读取 Excel，好处是读取方便，不足是一次性将文件加载到内存中，容易造成OOM；第二种模型：`eventusermodel`，这种方式采用事件驱动的方法解析 xml，在遇到文件内容时，事件会触发，这种做法可以大大降低内存的消耗。
