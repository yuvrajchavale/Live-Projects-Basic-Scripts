package july_5;

import java.io.FileInputStream;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile {

	public static void main(String[] args) throws Throwable {
		FileInputStream fi = new FileInputStream("D:\\Employee.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fi);
		XSSFSheet ws = wb.getSheet("Emp Data");
		XSSFRow rc = ws.getRow(0);
		int rowCount = ws.getLastRowNum();
		System.out.println("Total rows in Sheet: " + rowCount);
		int cellCount = rc.getLastCellNum();
		System.out.println("Total Cells in Sheet: " + cellCount);
		
		for (int i = 1; i <=rowCount; i++) {
			String fname = ws.getRow(i).getCell(0).getStringCellValue();
			String mname = ws.getRow(i).getCell(1).getStringCellValue();
			String lname = ws.getRow(i).getCell(2).getStringCellValue();
			int  eid = (int) ws.getRow(i).getCell(3).getNumericCellValue();
			System.out.println(fname +" "+mname +" "+lname +" "+eid);
			
		}
		wb.close();
		fi.close();
		
	

	}

}
