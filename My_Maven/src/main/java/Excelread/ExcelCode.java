package Excelread;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelCode {
	XSSFsheet sh;
	public ExcelCode() throws FileNotFoundException
	{
		FileInputStream f=new FileInputStream("C:\\Users\\Aswathyakhil\\eclipse-workspace\\My_Maven\\src\\main\\resources\\Book.xlsx");// for file open excelsheet
		
		XSSFWorkbook w= new XSSFWorkbook(f);
        sh= w.getSheet("Sheet1");
	
	}
public String readData(int row,int column) 
{
 Row r= sh.getRow(row);
Cell c=r.getCell(column);
int celltype=c.getCellType();
switch(celltype) 
{
case Cell.CELL_TYPE_NUMERIC;
{
	double val=c.getNumericCellValue();
	return String.copyValueOf(val);
	
}
case Cell.CELL_TYPE_STRING()
{
	return c.get
}
}
}

}