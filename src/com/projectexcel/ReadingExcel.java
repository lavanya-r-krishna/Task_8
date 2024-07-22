package com.projectexcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) {
		// path of an excel file
		String filepath="Test.xlsx";
		
		try {
			//create an input stream of fileinputstream object
			FileInputStream inpstrean=new FileInputStream(filepath);
			
			//create an object workbook
			XSSFWorkbook workbook=new XSSFWorkbook(inpstrean);
			
			//get an access to sheet
			Sheet sheet=workbook.getSheetAt(0);
			
			//iterate over the row in sheet
			for(Row row:sheet) {
				//iterate over the cells of each row
				for(Cell cell:row) {
					//check if the string or number
					if(cell.getCellType()==CellType.STRING) {
						System.out.print(cell.getStringCellValue()+"\t\t");
					}else if(cell.getCellType()==CellType.NUMERIC) {
						System.out.print(cell.getNumericCellValue()+"\t\t");
					}
					
				}
				System.out.println();
			}
			
		}catch(IOException e) {
			e.printStackTrace();
		}

	}

}
