package july_5;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateNewBook {

	public static void main(String[] args) throws Throwable {
		FileInputStream fi = new FileInputStream("D:\\Employee.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fi);
		XSSFSheet ws = wb.getSheet("Emp Data");
		XSSFRow row = ws.getRow(0);
		XSSFCell cell = row.getCell(0);
		int rc = ws.getLastRowNum();
		System.out.println(rc);
		for (int i = 1; i <=rc; i++) {
			String fname = ws.getRow(i).getCell(0).getStringCellValue();
			String mname = ws.getRow(i).getCell(1).getStringCellValue();
			String lname = ws.getRow(i).getCell(2).getStringCellValue();
			int eid = (int) ws.getRow(i).getCell(3).getNumericCellValue();
			//write passs into status cell
			ws.getRow(i).createCell(4).setCellValue("Pass");
			
			
		}
		
		fi.close();
		FileOutputStream fo = new FileOutputStream("D:/Updated.xlsx");
		wb.write(fo);
		
				

	}

}
