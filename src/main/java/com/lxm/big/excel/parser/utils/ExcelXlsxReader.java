package com.lxm.big.excel.parser.utils;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 
 * ClassName: com.lxm.big.excel.parser.utils.ExcelXlsxReader <br/>
 * Function: POI读取excel有两种模式，一种是用户模式，一种是事件驱动模式
 * 采用SAX事件驱动模式解决XLSX文件，可以有效解决用户模式内存溢出的问题， 该模式是POI官方推荐的读取大数据的模式，
 * 在用户模式下，数据量较大，Sheet较多，或者是有很多无用的空行的情况下，容易出现内存溢出
 * <p>
 * 用于解决.xlsx2007版本大数据量问题<br/>
 * Date: 2020年7月26日 下午3:08:56 <br/>
 * 
 * @author liuxiangming <br/>
 */
public class ExcelXlsxReader extends DefaultHandler {

	/**
	 * 单元格中的数据可能的数据类型
	 */
	enum CellDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
	}

	/**
	 * 共享字符串表
	 */
	private SharedStringsTable sst;

	/**
	 * 上一次的索引值
	 */
	private String lastIndex;

	/**
	 * 文件的绝对路径
	 */
	private String filePath = "";

	/**
	 * 工作表索引
	 */
	private int sheetIndex = 1;

	/**
	 * sheet名
	 */
	private String sheetName = "";

	/**
	 * 处理总行数
	 */
	private int totalRows = 0;

	/**
	 * 一行内cell集合
	 */
	private List<String> cellList = new ArrayList<String>();

	/**
	 * 判断整行是否为空行的标记
	 */
	private boolean flag = false;

	/**
	 * 当前列
	 */
	private int curCol = 0;

	/**
	 * T元素标识
	 */
	private boolean isTElement;

	/**
	 * 单元格数据类型，默认为字符串类型
	 */
	private CellDataType nextDataType = CellDataType.SSTINDEX;

	private final DataFormatter formatter = new DataFormatter();

	/**
	 * 单元格日期格式的索引
	 */
	private short formatIndex;

	/**
	 * 日期格式字符串
	 */
	private String formatString;

	// 定义前一个元素
	private String preRef = null;
	
	// 定义当前元素的位置（A1,B2等）
	private String ref = null;

	// 定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
	private String maxRef = null;

	// 定义当前元素名称
	private String tag;

	/**
	 * 单元格
	 */
	private StylesTable stylesTable;

	/**
	 * excel数据处理器
	 */
	private ExcelRowProcessor excelRowProcessor;

	/**
	 * 遍历工作簿中所有的电子表格，处理
	 * 
	 * @param filename
	 * @param excelRowProcessor
	 * @return
	 * @throws Exception
	 */
	public int process(String filename, ExcelRowProcessor excelRowProcessor)
			throws Exception {
		filePath = filename;
		this.excelRowProcessor = excelRowProcessor;
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader(pkg);
		stylesTable = xssfReader.getStylesTable();
		this.sst = xssfReader.getSharedStringsTable();
		XMLReader parser = XMLReaderFactory
				.createXMLReader("org.apache.xerces.parsers.SAXParser");
		parser.setContentHandler(this);
		XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader
				.getSheetsData();
		while (sheets.hasNext()) { // 遍历sheet
			InputStream sheet = sheets.next(); // sheets.next()和sheets.getSheetName()不能换位置，否则sheetName报错
			sheetName = sheets.getSheetName();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource); // 解析excel的每条记录，在这个过程中startElement()、characters()、endElement()这三个函数会依次执行
			sheet.close();
			this.sheetIndex++;
		}
		return totalRows; // 返回该excel文件的处理总行数
	}

	/**
	 * 第一个执行
	 *
	 * @param uri
	 * @param localName
	 * @param name
	 * @param attributes
	 * @throws SAXException
	 */
	@Override
	public void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		// c => 单元格
		if ("c".equals(name)) {
			// 前一个单元格的位置
			preRef = ref;
			// 当前单元格的位置
			ref = attributes.getValue("r");
			// 设定单元格类型
			this.setNextDataType(attributes);
		}
		// 当元素为t时
		if ("t".equals(name)) {
			isTElement = true;
		} else {
			isTElement = false;
		}
		// 置空
		lastIndex = "";
		// 设置当前元素名称
		tag = name;
	}

	/**
	 * 第二个执行 得到单元格对应的索引值或是内容值 如果单元格类型是字符串、INLINESTR、数字、日期，lastIndex则是索引值
	 * 如果单元格类型是布尔值、错误、公式，lastIndex则是内容值
	 * 
	 * @param ch
	 * @param start
	 * @param length
	 * @throws SAXException
	 */
	@Override
	public void characters(char[] ch, int start, int length)
			throws SAXException {
		if ("v".equals(tag)) { // v => 单元格的值，如果单元格是字符串，则v标签的值为该字符串在SST中的索引
			lastIndex += new String(ch, start, length);
		}
	}

	/**
	 * 第三个执行
	 *
	 * @param uri
	 * @param localName
	 * @param name
	 * @throws SAXException
	 */
	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		// t元素也包含字符串
		if (isTElement) {// 这个程序没经过
			// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
			String value = lastIndex.trim();
			cellList.add(curCol, value);
			curCol++;
			isTElement = false;
			// 如果里面某个单元格含有值，则标识该行不为空行
			if (value != null && !"".equals(value)) {
				flag = true;
			}
		} else if ("c".equals(name)) {
			// 补充该行空单元格
			int len = getColumnIndex(ref) - (preRef == null? 0: getColumnIndex(preRef)) - 1;
			for(int i = 0; i < len; i++){
				cellList.add(curCol, "");
				curCol++;
			}
			// v => 单元格的值，如果单元格是字符串，则v标签的值为该字符串在SST中的索引
			String value = this.getDataValue(lastIndex.trim());// 根据索引值获取对应的单元格值
			cellList.add(curCol, value);
			curCol++;
			// 如果里面某个单元格含有值，则标识该行不为空行
			if (value != null && !"".equals(value)) {
				flag = true;
			}
		} else {
			// 如果标签名称为row，这说明已到行尾，调用optRows()方法
			if ("row".equals(name)) {
				int curRow = getRowIndex(ref);
				// 默认第一行为表头，以该行单元格数目为最大数目
				if (curRow == 1) {
					maxRef = ref;
				}
				// 补全一行尾部可能缺失的单元格，避免数据处理时出现的越标问题
				if (maxRef != null) {
					int len = getColumnIndex(maxRef) - getColumnIndex(ref);
					for (int i = 0; i < len; i++) {
						cellList.add(curCol, "");
						curCol++;
					}
				}

				if (flag) { // 该行不为空行时进行处理
					excelRowProcessor.processRow(filePath, sheetName,
							sheetIndex, curRow, cellList);
					totalRows++;
				}

				cellList.clear();
				curCol = 0;
				preRef = null;
				ref = null;
				flag = false;
			}
		}
	}

	/**
	 * 处理数据类型
	 *
	 * @param attributes
	 */
	public void setNextDataType(Attributes attributes) {
		nextDataType = CellDataType.NUMBER; // cellType为空，则表示该单元格类型为数字
		formatIndex = -1;
		formatString = null;
		String cellType = attributes.getValue("t"); // 单元格类型
		String cellStyleStr = attributes.getValue("s"); //
//		String columnData = attributes.getValue("r"); // 获取单元格的位置，如A1,B1

		if ("b".equals(cellType)) { // 处理布尔值
			nextDataType = CellDataType.BOOL;
		} else if ("e".equals(cellType)) { // 处理错误
			nextDataType = CellDataType.ERROR;
		} else if ("inlineStr".equals(cellType)) {
			nextDataType = CellDataType.INLINESTR;
		} else if ("s".equals(cellType)) { // 处理字符串
			nextDataType = CellDataType.SSTINDEX;
		} else if ("str".equals(cellType)) {
			nextDataType = CellDataType.FORMULA;
		}

		if (cellStyleStr != null) { // 处理日期
			int styleIndex = Integer.parseInt(cellStyleStr);
			XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
			formatIndex = style.getDataFormat();
			formatString = style.getDataFormatString();
			if (formatString == null) {
                nextDataType = CellDataType.NULL;
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
			if (formatString.contains("m/d/yy")) {
				nextDataType = CellDataType.DATE;
				formatString = "yyyy-MM-dd hh:mm:ss";
			}
		}
	}

	/**
	 * 对解析出来的数据进行类型处理
	 * 
	 * @param value
	 *            单元格的值， value代表解析：BOOL的为0或1，
	 *            ERROR的为内容值，FORMULA的为内容值，INLINESTR的为索引值需转换为内容值，
	 *            SSTINDEX的为索引值需转换为内容值， NUMBER为内容值，DATE为内容值

	 * @return
	 */
	@SuppressWarnings("deprecation")
	public String getDataValue(String value) {
		if ("".equals(value) || value == null) {
			return "";
		}
		String thisStr = "";
		switch (nextDataType) {
		// 这几个的顺序不能随便交换，交换了很可能会导致数据错误
		case BOOL: // 布尔值
			char first = value.charAt(0);
			thisStr = first == '0' ? "FALSE" : "TRUE";
			break;
		case ERROR: // 错误
			thisStr = "\"ERROR:" + value + '"';
			break;
		case FORMULA: // 公式
			thisStr = '"' + value + '"';
			break;
		case INLINESTR:
			XSSFRichTextString rtsi = new XSSFRichTextString(value);
			thisStr = rtsi.toString();
			break;
		case SSTINDEX: // 字符串
			String sstIndex = value;
			try {
				int idx = Integer.parseInt(sstIndex);
				XSSFRichTextString rtss = new XSSFRichTextString(
						sst.getEntryAt(idx));// 根据idx索引值获取内容值
				thisStr = rtss.toString();
			} catch (NumberFormatException ex) {
				thisStr = value;
			}
			break;
		case NUMBER: // 数字
			if (formatString != null) {
				thisStr = formatter.formatRawCellContents(
						Double.parseDouble(value), formatIndex, formatString)
						.trim();
			} else {
				thisStr = value;
			}
			thisStr = thisStr.replace("_", "").trim();
			break;
		case DATE: // 日期
			thisStr = formatter.formatRawCellContents(
					Double.parseDouble(value), formatIndex, formatString);
			// 对日期字符串作特殊处理，去掉T
			thisStr = thisStr.replace("T", " ");
			break;
		default:
			thisStr = "";
			break;
		}
		return thisStr;
	}
	
    /**
     * 返回元素的列号（从1开始）
     * @param ref
     * @return
     */
    public int getColumnIndex(String ref){
    	String column = ref.replaceAll("\\d+", "");
    	int index = 0;
    	int pow = 1; // 初始为26的0次方，后续每循环一次乘以26一次
    	for(int i = column.length() - 1; i >= 0; i--){
    		index += (Character.toUpperCase(column.charAt(i)) - 64) *  pow;
    		pow *= 26;
    	}
    	return index;
    }
    
    /**
     * 返回当前元素的行号（从1开始）
     * @param ref
     * @return
     */
    public int getRowIndex(String ref){
    	return Integer.parseInt(ref.replaceAll("[A-Za-z]+", ""));
    }
    
}
