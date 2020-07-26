package com.lxm.big.excel.parser.utils;

import java.util.List;

/**
 * 
 * ClassName: com.lxm.big.excel.parser.utils.ExcelDataGenerator <br/>
 * Function: excel数据生成器，用于生成写入excel的数据<br/>
 * Date: 2020年7月26日 下午5:29:14 <br/>
 * @author liuxiangming <br/>
 */
public interface ExcelDataGenerator {
	
	/**
	 * 获取写入excel的数据
	 * @param startIndex 获取数据开始下标，从0开始
	 * @param size 一次获取数据大小
	 * @return
	 */
	public List<String[]> generateRowsData(int startIndex, int size);

}
