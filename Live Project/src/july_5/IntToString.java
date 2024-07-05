package july_5;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IntToString {

	public static void main(String[] args) throws Throwable {
		FileInputStream fi = new FileInputStream("D:\\Employee.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fi);
		XSSFSheet ws = wb.getSheet("Emp Data");
		XSSFRow row = ws.getRow(0);
		int rowCount = ws.getLastRowNum();
		XSSFCell cell = row.getCell(0);
		
		for (int i = 1; i <=rowCount; i++) {
			
			if (wb.getSheet("Emp Data").getRow(i).getCell(3).getCellType()==CellType.NUMERIC) {
				int cellData = (int) wb.getSheet("Emp Data").getRow(i).getCell(3).getNumericCellValue();
				String eid= String.valueOf(cellData);
				String fname = ws.getRow(i).getCell(0).getStringCellValue();
				String mname = ws.getRow(i).getCell(1).getStringCellValue();
				String lname = ws.getRow(i).getCell(2).getStringCellValue();
				System.out.println(fname + " " + mname+ " " + lname + " " + eid);
				ws.getRow(i).createCell(4).setCellValue("Failed");
				
			}
		}
		fi.close();
		FileOutputStream fo = new FileOutputStream("D:\\Int TO String.xlsx");
		wb.write(fo);
		fo.close();
		wb.close();

	}

}
