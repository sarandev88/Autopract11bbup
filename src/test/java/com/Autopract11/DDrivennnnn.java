package com.Autopract11;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DDrivennnnn {
	private static void readData() throws IOException {
		
	
	
		
		File f = new File ("D:\\saranya\\Autopract11\\src\\test\\java\\com\\Autopract11\\July timesheet.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
	 Workbook wb = new XSSFWorkbook(fis);
	 Sheet sheetAt = wb.getSheetAt(0);
	 
	 int row_size = sheetAt.getPhysicalNumberOfRows();
	 
	 //for loop
	 
	for(int i=0; i<row_size; i++) {
		Row row = sheetAt.getRow(i);
		int Cell_Size = row.getPhysicalNumberOfCells();
		for(int j=0; j<Cell_Size; j++) {
			

		 Cell cell = row.getCell(j);
		 CellType cellType = cell.getCellType();


		 
		 if(cellType.equals(cellType.STRING)) {
			 
			 String stringCellValue = cell.getStringCellValue();
			 
			 System.out.println(stringCellValue);
		 }
			 
			 else if(cellType.equals(cellType.NUMERIC)) {
				 
				 double numericCellValue = cell.getNumericCellValue();
				 
				 int value = (int) numericCellValue;
				 
				 System.out.println(value);
			 }
				 
				 
				 
				 
			 }
			 
		}
	}
			
	
		
public static void main(String[] args) throws Throwable {
	readData();
}
}


			
		
	
		
		 
		
	
		
		
		
		
		
	
	



