package com.lxm.big.excel.parser.utils;


/**
 * 
 * ClassName: com.lxm.big.excel.parser.utils.ExcelReaderUtil <br/>
 * Function: <br/>
 * Date: 2020年7月26日 下午3:13:37 <br/>
 * @author liuxiangming <br/>
 */
public class ExcelReaderUtil {
    //excel2003扩展名
    public static final String EXCEL03_EXTENSION = ".xls";
    //excel2007扩展名
    public static final String EXCEL07_EXTENSION = ".xlsx";

    public static int readExcel(String fileName, ExcelRowProcessor excelRowProcessor) throws Exception {
        int totalRows =0;
        if (fileName.toLowerCase().endsWith(EXCEL03_EXTENSION)) { //处理excel2003文件
            ExcelXlsReader excelXls=new ExcelXlsReader();
            totalRows =excelXls.process(fileName, excelRowProcessor);
        } else if (fileName.toLowerCase().endsWith(EXCEL07_EXTENSION)) {//处理excel2007文件
            ExcelXlsxReader excelXlsxReader = new ExcelXlsxReader();
            totalRows = excelXlsxReader.process(fileName, excelRowProcessor);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
        return totalRows;
    }

}