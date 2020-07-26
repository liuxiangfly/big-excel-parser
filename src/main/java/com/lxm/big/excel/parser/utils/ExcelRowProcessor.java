package com.lxm.big.excel.parser.utils;

import java.util.List;

/**
 * 
 * ClassName: com.lxm.big.excel.parser.utils.ExcelRowProcessor <br/>
 * Function: excel数据处理器 <br/>
 * Date: 2020年7月26日 下午4:10:33 <br/>
 * @author liuxiangming <br/>
 */
public interface ExcelRowProcessor {
	
	/**
	 * 处理一行数据
	 * @param filePath excel文件路径
	 * @param sheetName 表格名称
	 * @param sheetIndex 表格下标，从1开始
	 * @param curRow 当前行下标， 从1开始
	 * @param cellList 当前行数据
	 */
	public void processRow(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList);

}
