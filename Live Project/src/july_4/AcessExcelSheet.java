package july_4;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AcessExcelSheet {

	public static void main(String[] args) throws Throwable {
		
		//file input
		FileInputStream fi = new FileInputStream("D:\\Employee.xlsx");
		
		//Workbook from file path
		XSSFWorkbook wb = new XSSFWorkbook(fi);
		
		//Sheet from Workbook	
		XSSFSheet ws = wb.getSheet("Emp Data"); 
		
		//Aceess Row 
		XSSFRow row = ws.getRow(0);
		//Get row Count from sheet
		int rowCount = ws.getLastRowNum();
		System.out.println("Number of Rows: " + rowCount);
		
		//Acess cell 
		XSSFCell cell = row.getCell(0);
		
		int cellCount = row.getLastCellNum();
		System.out.println("Number of Cell: " + cellCount);
		
		
		String fname = ws.getRow(10).getCell(0).getStringCellValue();
		String mname = ws.getRow(10).getCell(1).getStringCellValue();
		String lname = ws.getRow(10).getCell(2).getStringCellValue();
		int empId = (int) ws.getRow(10).getCell(3).getNumericCellValue();
		
		System.out.println(fname +"\n"+ mname+"\n"+lname+"\n"+empId);

	}

}
