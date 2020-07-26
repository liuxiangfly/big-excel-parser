package com.lxm.big.excel.parser;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.lxm.big.excel.parser.utils.ExcelDataGenerator;
import com.lxm.big.excel.parser.utils.ExcelReaderUtil;
import com.lxm.big.excel.parser.utils.ExcelRowProcessor;
import com.lxm.big.excel.parser.utils.ExcelWriterUtil;


public class Main {

    public static void main(String[] args) throws Exception {
    	write();
    }
    
    private static void write(){
    	String path="D:\\Documents\\excel数据分析\\数据表(1)\\数据表\\test.xlsx";
    	ExcelDataGenerator generator = new ExcelDataGenerator() {
			
			@Override
			public List<String[]> generateRowsData(int startIndex, int size) {
				if(startIndex > 9320){
					return null;
				}
				int len = Math.min(9320, startIndex + size);
				List<String[]> rows = new ArrayList<>();
				for(int i = startIndex; i < len; i++){
					String[] row = new String[20];
					for(int j = 0; j < 20; j++){
						row[j] = "data" + i + "-" + j;
					}
					rows.add(row);
				}
				return rows;
			}
		};
		ExcelWriterUtil.generateExcel(path, "生成结果", generator);
    }
    
    private static void read() throws Exception{
    	String path="D:\\Documents\\excel数据分析\\数据表(1)\\数据表\\原始数据\\江苏华为医药物流有限公司.xlsx";
        ExcelRowProcessor excelRowProcessor = new ExcelRowProcessor() {
			
			@Override
			public void processRow(String filePath, String sheetName, int sheetIndex,
					int curRow, List<String> cellList) {
				StringBuilder oneLineSb = new StringBuilder();
	            oneLineSb.append(filePath);
	            oneLineSb.append("--");
	            oneLineSb.append("sheet" + sheetIndex);
	            oneLineSb.append("::" + sheetName);//加上sheet名
	            oneLineSb.append("--");
	            oneLineSb.append("row" + curRow);
	            oneLineSb.append("::");
	            for (String cell : cellList) {
	                oneLineSb.append(cell.trim());
	                oneLineSb.append("|");
	            }
	            String oneLine = oneLineSb.toString();
	            if (oneLine.endsWith("|")) {
	                oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
	            }// 去除最后一个分隔符

	            System.out.println(oneLine);
				
			}
		};
        ExcelReaderUtil.readExcel(path, excelRowProcessor);
    }

}
