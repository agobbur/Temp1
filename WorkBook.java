package com.ag.utility;

import java.io.*;
import java.util.*;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

public class WorkBook {
   public static void main(String[] args)throws Exception {
	   	//Create blank workbook
	   	XSSFWorkbook workbookW1 = new XSSFWorkbook(); 
	   
		//Create blank spreadsheet in the workbook
		XSSFSheet spreadsheetW1 = workbookW1.createSheet("RateCard");
		XSSFSheet spreadsheetW2 = workbookW1.createSheet("Instructions");
		
		//Create row
		XSSFRow row;

		//Prepare data to write in cells
		Map <String, Object[]> rateCard = new TreeMap <String, Object[]>();
		rateCard.put( "1", new Object[] {"VENDO_ID", "PRODUCT_ID", "PRODUCT_NAME", "TERM", "MRR"});
		rateCard.put( "2", new Object[] {"99",	     "1",          "T1 1.5",       "1",    "200.00"});
		rateCard.put( "3", new Object[] {"99",	     "2",	   	   "T1 3.0",	   "1",	   "300.00"});

		//Iterate over data and write to spreadsheet
		Set<String> keyid = rateCard.keySet();
		int rowid = 0;
		for (String key:keyid){
			row = spreadsheetW1.createRow(rowid++);
			Object [] objectArr = rateCard.get(key);
			int cellid = 0;
			for (Object obj:objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String)obj);
			}
		}
		
		//Create output stream to write to specified file
		FileOutputStream fos = new FileOutputStream(new File("RateCard_Vendor-99.xlsx"));
		
		//Write to workbook using file-out object 
		workbookW1.write(fos);
		System.out.println("Excel workbook (xlsx) created successfully, with spreadsheets and cell data in it.");
		
		fos.close();
		workbookW1.close();
	  
	  	//----------------------------------------------------------------------
	  	//Open an existing specified workbook
	  	
	  	//Create input stream to read specified file
		File file = new File("RateCard_Vendor-99.xlsx");
		FileInputStream fis = new FileInputStream(file);
	 	
		//Get the workbook instance for XLSX file by passing input-stream
		XSSFWorkbook workbookR1 = new XSSFWorkbook(fis);
		
		if(file.isFile() && file.exists()) {
			System.out.println("Workbook opened successfully.");
			workbookR1.close();
		}
		else {
			System.out.println("Error: Could not open the Workbook.");
		}
		
		//----------------------------------------------------------------------
		//Read the opened workbook
		
		//Read first spreadsheet and output the values read from cells to console
		System.out.println("Reading cell values in the first spreadsheet of the workbook...");
		XSSFSheet spreadsheetR1 = workbookR1.getSheetAt(0);
		
		Iterator <Row> rowIterator = spreadsheetR1.iterator();
		
		while (rowIterator.hasNext()) {
			row = (XSSFRow) rowIterator.next();
			Iterator <Cell> cellIterator = row.cellIterator();
			
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + ", ");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + ", ");
						break;
				}
			}
			System.out.println();
		}
		System.out.println("Done.");
		fis.close();
   }
}