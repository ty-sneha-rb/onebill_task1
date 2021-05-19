package com.ty.excelColName;

public class ReadExcelDataUsingUtilClass {

	public static void main(String[] args) throws Exception{
		ExcelApiTest eat = new ExcelApiTest("./data/task.xlsx.xltx");
		System.out.println(eat.getCellData("Sheet1", 0, 1));
		System.out.println(eat.getCellData("Sheet1", 1, 1));
		System.out.println(eat.getCellData("Sheet1", 2, 1));
		
		System.out.println("***********************");
		
		System.out.println(eat.getCellData("Sheet1", "Product_Name", 1));
		System.out.println(eat.getCellData("Sheet1", "Onetime_Fee", 2));
		System.out.println(eat.getCellData("Sheet1", "Recurring_Fee", 3));
		
		
	}

}
