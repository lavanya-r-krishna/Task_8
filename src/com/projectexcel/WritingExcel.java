package com.projectexcel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		//create an blank excel sheet of xssfworkbook
		//try with resouces
		
		
		try(XSSFWorkbook workbook=new XSSFWorkbook()){
			//create sheet
			
			Sheet sheet=workbook.createSheet("sheet1");
			
			//create an array of objects
			Object[][] data= {
					{"Name","Age","Email"},
					{"John Doe","30","john@test.com"},
					{"Jane Doe","28","john@test.com"},
					{"Bob Smith","35","jacky@example.com"},
					{"Swapnil","37","swapnil@example.com"},
					
			};
			
			//writing the data into excel
			int rowNum=0;
			for(Object[] rowdata:data) {
				
				//create a row in the sheet
				Row row=sheet.createRow(rowNum++);
				
				//insert data into cells
				int colNum=0;// to print colnum
				for(Object field:rowdata) {
					Cell cell=row.createCell(colNum++);//created cell
					if(field instanceof String) {
						cell.setCellValue((String)field);
					}else if(field instanceof Integer) {
						cell.setCellValue((Integer)field);
					}
				}
			}
			
			//create a file outstream object and write the data
			try(FileOutputStream os=new FileOutputStream("Test.xlsx")){
				workbook.write(os);
			}
			System.out.println("Data Added Successfully to file..");
		}catch (IOException e) {
			e.printStackTrace();
	    }

   }
}

