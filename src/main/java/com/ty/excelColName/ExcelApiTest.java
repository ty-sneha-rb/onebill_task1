package com.ty.excelColName;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class ExcelApiTest {

	public FileInputStream fis = null;
	public FileOutputStream fos = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	String xlFilePath;

	public ExcelApiTest(String xlFilePath) throws Exception{
		this.xlFilePath = xlFilePath;
		fis= new FileInputStream(xlFilePath);
		workbook = new XSSFWorkbook(fis);
		fis.close();
	}
	public String getCellData(String sheetName , int colNum , int rowNum) {
		try {
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);

			if(cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.STRING)
				return cell.getStringCellValue();
			else if(cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.NUMERIC || cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.FORMULA)
			{
				String cellValue = String.valueOf(cell.getNumericCellValue());
				if(HSSFDateUtil.isCellDateFormatted(cell)) {
					DateFormat dt = new SimpleDateFormat("dd/mm/yy");
					Date date = cell.getDateCellValue();
					cellValue = dt.format(date);	
				}
				return cellValue ;
			}
			else if(cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		}
		catch(Exception e)
		{
			e.printStackTrace();
			return "NOT MATCHTED VALUE";
		}
	}
	
	public String getCellData(String sheetName , String colName, int rowNum) {
		try {
			int colNum = 0;
			sheet = workbook.getSheet(sheetName);
			row= sheet.getRow(0);
			
			for(int i=0; i<row.getLastCellNum(); i++) {
				if(row.getCell(i).getStringCellValue().trim().equals("colName"))
					colNum= i;
			}
			row= sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			
			if(cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.STRING)
				return cell.getStringCellValue();
			else if(cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.NUMERIC || cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.FORMULA)
			{
				String cellValue = String.valueOf(cell.getNumericCellValue());
				if(HSSFDateUtil.isCellDateFormatted(cell)) {
					DateFormat dt = new SimpleDateFormat("dd/mm/yy");
					Date date = cell.getDateCellValue();
					cellValue = dt.format(date);	
				}
				return cellValue ;
			}
			else if(cell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
			
		}catch(Exception e)
		{
			e.printStackTrace();
			return "data not found";
		}
		
	}
}
