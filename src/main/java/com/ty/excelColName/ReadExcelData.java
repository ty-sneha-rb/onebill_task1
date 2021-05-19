package com.ty.excelColName;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {
	
public static void main(String[] args) throws Exception {
	FileInputStream fis = new FileInputStream("./data/task.xlsx.xltx");
	XSSFWorkbook workbook = new XSSFWorkbook(fis);
	XSSFSheet sheet = workbook.getSheet("Sheet1");
	XSSFRow row = sheet.getRow(0);
	XSSFCell cell = null;
	
	int colNum = -1;
	for(int i=0; i<row.getLastCellNum(); i++) {
		if(row.getCell(i).getStringCellValue().trim().equals("Product_Name"))
			colNum= i;
	}
	
	row= sheet.getRow(1);
	cell = row.getCell(colNum);
	
	String product_Name = String.valueOf(cell.getStringCellValue());
	System.out.println("value from the excel sheet :"+ product_Name);
}
}
