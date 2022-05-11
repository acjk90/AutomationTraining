package Excelread;

import java.io.IOException;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		 
		ExcelCode obj = new ExcelCode();
		
		String str = obj.readData(0, 0);		
		System.out.println(str);
		
		String str1 = obj.readData(0, 1);		
		System.out.println(str1);
		
		String str2 = obj.readData(0, 2);		
		System.out.println(str2);
		
		System.out.println();

				
		String str3 = obj.readData(1, 0);		
		System.out.println(str3);
		
		String str4 = obj.readData(1, 1);		
		System.out.println(str4);
		
		String str5 = obj.readData(1, 2);		
		System.out.println(str5);		

		String str6 = obj.readData(0, 3);		
		System.out.println(str6);
		
	}

}

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode 
{
	XSSFSheet sh;
	
	public ExcelCode() throws IOException
	{
		FileInputStream f = new FileInputStream("C:\\Users\\ARCHANA\\eclipse-workspace\\My_Maven\\src\\main\\resources\\Book.xlsx.xlsx"); //File Open
		XSSFWorkbook w = new XSSFWorkbook(f);
		sh = w.getSheet("Sheet1");
		           
	}
	
	public String readData(int row,int column)
	{
		Row r = sh.getRow(row);
		Cell c = r.getCell(column);
		int celltype = c.getCellType();
		switch(celltype) 
		{
		case Cell.CELL_TYPE_NUMERIC:
		{
			double val = c.getNumericCellValue();
			return String.valueOf(val);
		}
		case Cell.CELL_TYPE_STRING:
		{
			return c.getStringCellValue();
		}
		}
		
		return null;
	}
}