package com.Cyspan.training;


	

	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.IOException;

	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class ExcelCode {
		public static FileInputStream f;
		public static XSSFWorkbook wb;
		public static XSSFSheet sh;
		
		
		public static String readStringData(int i,int j) throws IOException
		{
			f=new FileInputStream("D:\\Maven sample\\TestExcel.xlsx");
			wb=new XSSFWorkbook(f);
			sh=wb.getSheet("Sheet1");//sheet name
			Row r=sh.getRow(i);
			Cell c=r.getCell(j);
			return c.getStringCellValue();
		}
		
		public static double readNumericData(int i,int j) throws IOException
		{
			f=new FileInputStream("D:\\Maven sample\\TestExcel.xlsx");
			wb=new XSSFWorkbook(f);
			sh=wb.getSheet("Sheet1");//sheet name
			Row r=sh.getRow(i);
			Cell c=r.getCell(j);
			return c.getNumericCellValue();
		}
		

	}



