package com.lxm.big.excel.parser.utils;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * 
 * ClassName: com.lxm.big.excel.parser.utils.ExcelWriter <br/>
 * Function: excel写入器 <br/>
 * Date: 2020年7月26日 下午5:19:58 <br/>
 * @author liuxiangming <br/>
 */
public class ExcelWriter {

	private SXSSFWorkbook wb;

	private Sheet sh;

	private String filePath;
	
	private int rowIndex = 0;

	/**
	 * 创建电子表格
	 * @param filePath
	 * @param sheetName
	 */
	public void createExcel(String filePath, String sheetName) {
		this.filePath = filePath;
		// keep 100 rows in memory, exceeding rows will be flushed to disk
		this.wb = new SXSSFWorkbook(100); 
		this.sh = this.wb.createSheet(sheetName);
	}

	/**
	 * 将数据写入电子表格
	 * @param rows
	 */
	public void write(List<String[]> rows) {
		if(rows == null || rows.isEmpty()){
			return;
		}
		for (int i = 0; i < rows.size(); i++) {
			String[] rowData = rows.get(i);
			Row row = this.sh.createRow(rowIndex++);
			for (int j = 0; j < rowData.length; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue(rowData[j]);
			}
		}
	}

	/**
	 * 完成电子表格写入：将电子表格数据写入文件，同时处理掉产生的临时备份文件数据
	 */
	public void finish() {
		try (FileOutputStream out = new FileOutputStream(this.filePath)) {
			((SXSSFSheet) sh).flushRows(); // flush all others
			wb.write(out);
			// dispose of temporary files backing this workbook on disk
			wb.dispose();
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

}
