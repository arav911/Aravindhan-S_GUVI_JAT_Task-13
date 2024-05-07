package task_13_guvi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromExcel {

	public static void main(String[] args) throws IOException {

//		Specify the location of file
		File file = new File(".\\src\\task_13_guvi\\Aravindhan-S_GUVI_JAT_Task-13.xlsx");
		
//		Load file
		FileInputStream fis = new FileInputStream(file);
		
//		Load workbook
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
//		Load work sheet
		XSSFSheet sheet = wb.getSheet("Sheet1");
		
//		get total number of rows
		int rows = sheet.getPhysicalNumberOfRows();
		
//		get total number of columns
		int columns = sheet.getRow(0).getPhysicalNumberOfCells();
		
//		print all the values from the excel
		for(int i=0; i<rows; i++) {
			for(int j=0; j<columns; j++) {
				if(sheet.getRow(i).getCell(j).getCellType().equals(CellType.STRING))
					System.out.print(sheet.getRow(i).getCell(j).getStringCellValue()+" ");
				if(sheet.getRow(i).getCell(j).getCellType().equals(CellType.NUMERIC))
					System.out.print((int)sheet.getRow(i).getCell(j).getNumericCellValue()+" ");
			}
			System.out.println();
		}
		
		wb.close();

	}

}
