package task_13_guvi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataToExcel {

	public static void main(String[] args) throws IOException {
//		create object of workbook
		XSSFWorkbook wb = new XSSFWorkbook();
		
//		create object of work sheet
		XSSFSheet sheet = wb.createSheet("Sheet1");
		
//		Create ArrayList
		ArrayList<Object[]> studentDetails = new ArrayList<Object[]>();
		studentDetails.add(new Object[] {"Roll Number", "Name", "Department"});
		studentDetails.add(new Object[] {1001, "Rajesh", "IT"});
		studentDetails.add(new Object[] {1002, "Kumar", "ECE"});
		studentDetails.add(new Object[] {1003, "Scofield", "CSE"});
		studentDetails.add(new Object[] {1004, "Matthew", "EEE"});
		
		int rowNum = 0;
//		outer loop for rows
		for(Object[] student: studentDetails) {
			XSSFRow row = sheet.createRow(rowNum++);
			int cellNum = 0;
//			inner loop for columns
			for(Object stud: student) {
				XSSFCell cell = row.createCell(cellNum++);
				if(stud instanceof String)
					cell.setCellValue((String) stud);
				if(stud instanceof Integer)
					cell.setCellValue((Integer) stud);
				if(stud instanceof Boolean)
					cell.setCellValue((Boolean) stud);
			}
		}
//		Give filepath where StudentDetails.xlsx will create
		String filePath = ".\\src\\task_13_guvi\\StudentDetails.xlsx";
//		create object of FileOutputStream
		FileOutputStream fos = new FileOutputStream(filePath);
//		write data to excel
		wb.write(fos);
//		close FileOutputStream
		fos.close();
		wb.close();
		System.out.println("Data written to Excel file successfully");
	}

}
