package com.lxm.big.excel.parser.utils;

import java.util.List;

/**
 * 
 * ClassName: com.lxm.big.excel.parser.utils.ExcelWriterUtil <br/>
 * Function: 生成excel工具 <br/>
 * Date: 2020年7月26日 下午5:36:03 <br/>
 * @author liuxiangming <br/>
 */
public class ExcelWriterUtil {
	
	/**
	 * 生成excel
	 * @param filePath excel文件路径
	 * @param sheetName excel表格名
	 * @param excelDataGenerator excel数据获取器
	 */
	public static void generateExcel(String filePath, String sheetName, ExcelDataGenerator excelDataGenerator){
		ExcelWriter writer = new ExcelWriter();
		writer.createExcel(filePath, sheetName);
		int startIndex = 0;
		int size = 300;
		List<String[]> rows = null;
		do{
			rows = excelDataGenerator.generateRowsData(startIndex, size);
			writer.write(rows);
			startIndex += size;
		}while(rows != null && rows.size() >= size);
		writer.finish();
	}

}
