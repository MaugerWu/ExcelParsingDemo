package com.cqupt.mauger.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Java 解析 Excel（xls、xlsx）
 * @author Mauger
 * @date 2018年9月11日  
 * @version 1.0
 */
public class ExcelParsingTest
{
	
	/* Excel文件扩展名 xls、xlsx */
	public static final String EXCEL_FILE_EXTENSION_XLS = ".xls"; 
	public static final String EXCEL_FILE_EXTENSION_XLSX = ".xlsx"; 
	
	/* 解决POI在导入数字时，获取Cell值已经损失精度问题 */
	private static NumberFormat numberFormat = NumberFormat.getInstance();
	static
	{
		// 设为 false，指不使用分组方式显示数据，打印 9999999；设为 true，则打印 9,999,999
		numberFormat.setGroupingUsed(false);
	}
	
	public static void main(String[] args)
	{
		Workbook wb = null; // 对应 Excel 文档
		Sheet sheet = null; // 对应 Excel 文档中的一个 Sheet
		Row row = null; // 对应 Excel 文档中的一个 Sheet 中的一行
		List<Map<String, String>> list = null; // 用来存放表中的数据
		String cellData = null;
		
		String filePath = "src/Test.xlsx";
		wb = readExcel2(filePath);
		if (null != wb)
		{
			list = new ArrayList<Map<String,String>>();
			// 获取第一个 Sheet
			sheet = wb.getSheetAt(0);
			// 获取最大的行数
			int rownum = sheet.getPhysicalNumberOfRows();
			// 获取第一行
			row = sheet.getRow(0);
			// 获取最大的列数
			int colnum = row.getPhysicalNumberOfCells();
			// 记录目录
			String[] columns = new String[colnum];
			
			for (int i = 0; i < rownum; i++)
			{
				Map<String, String> map = new LinkedHashMap<String, String>();
				row = sheet.getRow(i);
				if (null != row)
				{
					if (i == 0)
					{
						for (int j = 0; j < colnum; j++)
						{
							columns[j] = getCellFormatValue(row.getCell(j));
						}
					} else
					{
						for (int j = 0; j < colnum; j++)
						{
							cellData = getCellFormatValue(row.getCell(j));
							map.put(columns[j], cellData);
						}
					}
				} else
				{
					break;
				}
				list.add(map);
			}
			
			// 遍历解析出来的 list
			for (Map<String, String> map : list)
			{
				for (Entry<String, String> entry : map.entrySet())
				{
					System.out.print(entry.getKey() + ": " + entry.getValue().toString() + ", ");
				}
				System.out.println();
			}
		}
	}

	/**
	 * 读取 Excel 文档
	 * 在没有使用 poi-ooxml-3.8.jar 中的 WorkbookFactory 类时，需要判断文档的扩展名 xls/xlsx
	 * @param filePath File Path
	 * @return Workbook
	 */
	@SuppressWarnings("resource")
	private static Workbook readExcel(String filePath)
	{
		if (null == filePath)
		{
			return null;
		}
		
		Workbook wb = null;
		// 获取 Excel 文档的扩展名 xls/xlsx
		String extString = filePath.substring(filePath.lastIndexOf("."));
		InputStream is = null;
		try
		{
			System.out.println("Parsing... " + filePath);
			is = new FileInputStream(filePath);
			if (EXCEL_FILE_EXTENSION_XLS.equals(extString))
			{
				try
				{
					return wb = new XSSFWorkbook(is);
				} catch (IOException e)
				{
					e.printStackTrace();
				}
			} else if (EXCEL_FILE_EXTENSION_XLSX.equals(extString))
			{
				try
				{
					return wb = new XSSFWorkbook(is);
				} catch (IOException e)
				{
					e.printStackTrace();
				}
			} else
			{
				return wb = null;
			}
		} catch (FileNotFoundException e)
		{
			e.printStackTrace();
		}
		
		return wb;
	}
	
	/**
	 * 这里主要是使用 POI 提供的 poi-ooxml-3.8.jar 中的 WorkbookFactory 类
	 * 只需要简单的几行代码就可以实例化 Workbook，且不用管它是 .xsl 或 .xslx
	 * @param filePath File Path
	 * @return Workbook
	 */
	private static Workbook readExcel2(String filePath)
	{
		if (null == filePath)
		{
			return null;
		}
		
		InputStream is = null;
		Workbook wb = null;
		try
		{
			is = new FileInputStream(filePath);
			wb = WorkbookFactory.create(is);
		} catch (FileNotFoundException e)
		{
			e.printStackTrace();
		} catch (InvalidFormatException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
		
		return wb;
	}
	
	/**
	 * 解析列
	 * @param cell Cell
	 * @return Object
	 */
	private static String getCellFormatValue(Cell cell)
	{
		String cellValue = null;
		if (null != cell)
		{
			// 判断 cell 的类型
			switch (cell.getCellType())
			{
			case Cell.CELL_TYPE_NUMERIC: // 数值
			{
				// 判断 cell 是否为日期格式
				if (DateUtil.isCellDateFormatted(cell))
				{
					// 转化为日期格式：YYYY-MM-DD
					cellValue = cell.getDateCellValue().toString();
				} else
				{
					// 数字
					cellValue = numberFormat.format(cell.getNumericCellValue());
				}
				break;
			}
			case Cell.CELL_TYPE_FORMULA: // 公式
			{
				cellValue = String.valueOf(cell.getCellFormula());
				break;
			}
			case Cell.CELL_TYPE_STRING:
			{
				cellValue = cell.getRichStringCellValue().toString();
				break;
			}
			default:
				cellValue = "";
			}
		} else
		{
			cellValue = "";
		}
		
		return cellValue;
	}
}
