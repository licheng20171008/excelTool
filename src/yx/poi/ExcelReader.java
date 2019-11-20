package yx.poi;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	// 创建文件输出流
	private BufferedReader reader = null;
	
	// 文件类型
	private String fileType;
	
	// 文件二进制输入流
	private InputStream is = null;
	
	// 当前sheet
	private int currSheet;
	
	// 当前位置
	private int currPosition;
	
	// sheet数量
	private int numOfSheets;
	
	// HSSFWorkbook
	Workbook workbook = null;
	
	// 设置cell之间以空格分割
	private static String EXCEL_LINE_DELIMITER = " ";
	
	// 设置最大列数
//	private static int MAX_EXCEL_COLUMNS = 64;
	
	public ExcelReader(String inputfile) throws IOException {
		
		// 判断参数是否为空或没有意义
		if (inputfile == null || "".equals(inputfile.trim())) {
			throw new IOException("no input file specified");
		}
		
		// 取得文件名的后缀名赋值给filetype
		this.fileType = inputfile.substring(inputfile.lastIndexOf(".") + 1);
		
		// 设置开始为0
		currPosition = 0;
		
		// 设置当前位置为0
		currPosition = 0;
		
		// 创建文件输入流
		is = new FileInputStream(inputfile);
		
		// 判断文件格式
		if (fileType.equalsIgnoreCase("txt")) {
			
			// 如果是TXT则直接创建BufferedReader读取
			reader = new BufferedReader(new InputStreamReader(is));
		} else if (fileType.equalsIgnoreCase("xls")) {
			
			// 如果是EXCEL文件则创建HSSFWorkbook读取
			workbook = new HSSFWorkbook(is);
			
			// 设置Sheet数
			numOfSheets = workbook.getNumberOfSheets();
		} else if (fileType.equalsIgnoreCase("xlsx")) {
			
			// 如果是EXCEL文件则创建HSSFWorkbook读取
			workbook = new XSSFWorkbook(is);
						
			// 设置Sheet数
			numOfSheets = workbook.getNumberOfSheets();
		} 
		else {
			throw new IOException("File Type Not Supported");
		}
	}
	
	// 函数readLine读取文件的一行
	public String readLine() throws IOException {
		
		// 如果是txt文件则通过reader读取
		if (fileType.equalsIgnoreCase("txt")) {
			
			String str = reader.readLine();
			
			// 空行则略去，直接读取下一行
			while ("".equals(str.trim())) {
				str = reader.readLine();
			}
			return str;
		} else if (fileType.equalsIgnoreCase("xls")) {
			
			// 如果是xls文件则通过POI提供的API读取文件
			// 根据currSheet值获得当前的SHEET
			HSSFSheet sheet = (HSSFSheet) workbook.getSheetAt(currSheet);
			
			// 判断当前行是否到当前sheet的结尾
			if (currPosition > sheet.getLastRowNum()) {
				
				// 当前行位置清零
				currPosition = 0;
				
				// 判断是否还有Sheet
				while(currSheet != numOfSheets -1) {
					
					// 得到下一张Sheet
					sheet = (HSSFSheet) workbook.getSheetAt(currSheet + 1);
					
					// 当前行数是否已经到达文件末尾
					if (currPosition == sheet.getLastRowNum()) {
						
						// 当前SHEET指向下一张sheet
						currSheet++;
						continue;
					}else {
						// 获取当前行数
						int row = currPosition;
						currPosition++;
						
						// 读取当前行数据
						return getHssfSheetLine(sheet, row);
					}
				}
				return null;
			}
			
			// 获取当前行数
			int row = currPosition;
			currPosition++;
			
			// 读取当前行数据
			return getHssfSheetLine(sheet, row);
		} else if (fileType.equalsIgnoreCase("xlsx")) {
			
			// 如果是xls文件则通过POI提供的API读取文件
			// 根据currSheet值获得当前的SHEET
			XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(currSheet);
			
			// 判断当前行是否到当前sheet的结尾
			if (currPosition > sheet.getLastRowNum()) {
				
				// 当前行位置清零
				currPosition = 0;
				
				// 判断是否还有Sheet
				while(currSheet != numOfSheets -1) {
					
					// 得到下一张Sheet
					sheet = (XSSFSheet) workbook.getSheetAt(currSheet + 1);
					
					// 当前行数是否已经到达文件末尾
					if (currPosition == sheet.getLastRowNum()) {
						
						// 当前SHEET指向下一张sheet
						currSheet++;
						continue;
					}else {
						// 获取当前行数
						int row = currPosition;
						currPosition++;
						
						// 读取当前行数据
						return getXssfSheetLine(sheet, row);
					}
				}
				return null;
			}
			
			// 获取当前行数
			int row = currPosition;
			currPosition++;
			
			// 读取当前行数据
			return getXssfSheetLine(sheet, row);
		}
		return null;
	}
	
	// 函数getLine返回Sheet的一行数据
	private String getHssfSheetLine(HSSFSheet sheet, int row) {
		
		// 根据行数取得sheet的一行
		HSSFRow rowline = sheet.getRow(row);
		
		// 创建字符创缓冲区
		StringBuffer buffer = new StringBuffer();
		
		// 获取当前行的列数
		int filledColumns = 0;
		if (rowline != null) {
			filledColumns = rowline.getLastCellNum();
		}
		HSSFCell cell = null;
		
		// 循环遍历所有列
		for (int i = 0; i < filledColumns; i++) {
			
			// 取得当前cell
			cell = rowline.getCell(i);
			String cellvalue = null;
			if (cell != null) {
				// 判断当前cell的type
				switch(cell.getCellType()){
				// 如果当前cell的type为NUMERIC
				case HSSFCell.CELL_TYPE_NUMERIC:{
					// 判断当前cell是否为date
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						
						// 把date转换成本地格式的字符串
						cellvalue = cell.getDateCellValue().toString();
					} else {
						// 如果是纯数字
						// 取得当前cell的数值
						Integer num = new Integer((int) cell.getNumericCellValue());
						cellvalue = String.valueOf(num);
					}
					break;
					
				// 如果当前cell的type为STRING
				} case HSSFCell.CELL_TYPE_STRING:{
					// 取得当前的cell字符串
					cellvalue = cell.getStringCellValue().replaceAll("'", "");
					break;
				} default:{
					cellvalue = " ";
				}
				}
			} else {
				cellvalue = "";
			}
			
			// 在每个字段之间插入分隔符
			buffer.append(cellvalue).append(EXCEL_LINE_DELIMITER);
		}
		
		// 以字符串返回该行的数据
		return buffer.toString();
	}

	// 函数getLine返回Sheet的一行数据
		private String getXssfSheetLine(XSSFSheet sheet, int row) {
			
			// 根据行数取得sheet的一行
			XSSFRow rowline = sheet.getRow(row);
			
			// 创建字符创缓冲区
			StringBuffer buffer = new StringBuffer();
			
			// 获取当前行的列数
			int filledColumns = 0;
			if (rowline != null) {
				filledColumns = rowline.getLastCellNum();
			}
			XSSFCell cell = null;
			
			// 循环遍历所有列
			for (int i = 0; i < filledColumns; i++) {
				
				// 取得当前cell
				cell = rowline.getCell(i);
				String cellvalue = null;
				if (cell != null) {
					// 判断当前cell的type
					switch(cell.getCellType()){
					// 如果当前cell的type为NUMERIC
					case XSSFCell.CELL_TYPE_NUMERIC:{
						// 判断当前cell是否为date
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							
							// 把date转换成本地格式的字符串
							cellvalue = cell.getDateCellValue().toString();
						} else {
							// 如果是纯数字
							// 取得当前cell的数值
							Integer num = new Integer((int) cell.getNumericCellValue());
							cellvalue = String.valueOf(num);
						}
						break;
						
					// 如果当前cell的type为STRING
					} case HSSFCell.CELL_TYPE_STRING:{
						// 取得当前的cell字符串
						cellvalue = cell.getStringCellValue().replaceAll("'", "");
						break;
					} default:{
						cellvalue = " ";
					}
					}
				} else {
					cellvalue = "";
				}
				
				// 在每个字段之间插入分隔符
				buffer.append(cellvalue).append(EXCEL_LINE_DELIMITER);
			}
			
			// 以字符串返回该行的数据
			return buffer.toString();
		}
		
	// close函数执行流的关闭操作
	public void close(){
		// 如果IS为空，则关闭InputStream文件输入流
		if (is != null) {
			try {
				is.close();
			} catch (IOException e) {
				is = null;
			}
		}
		
		// 如果reader不为空则关闭BufferedReader文件输入流
		if (reader != null) {
			try {
				reader.close();
			} catch (IOException e) {
				reader = null;
			}
		}
	}
	public static void main(String[] args) {
		try {
			ExcelReader er = new ExcelReader("E:\\流程导出\\14楼在建工程材料.xlsx");
			String line = er.readLine();
			while (line != null) {
				System.out.println(line);
				line = er.readLine();
			}
			er.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
