package com.tyss.JavaFileHandling;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static List<String> readExcelFile() {
		ArrayList<String> list = new ArrayList<String>();
		try {

			FileInputStream file = new FileInputStream(new File("Java.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			DataFormatter formatter = new DataFormatter();


			// Iterate through each rows one by one

			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {

				String object = "";
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();
					String cellValue = formatter.formatCellValue(cell);
					// Check the cell type and format accordingly
				
					object +=cellValue;

				}
				list.add(object);
			}
			file.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;

	}

	public static void writeExcel(List<String> list) {
			
		ArrayList<Integer> list1=new ArrayList<Integer>();
		for(int i=0; i<list.size(); i++) {
			list1.add(i);
		}
		Collections.shuffle(list);
		

		// Using XSSF for xlsx format, for xls use HSSF
		Workbook workbook = new XSSFWorkbook();

		Sheet questionSheet = workbook.createSheet("JavaSheet");

		int rowIndex = 0;
		
		for(int i=0; i<40; i++) {
			int cellIndex = 0;
			Row row = questionSheet.createRow(rowIndex++);
			row.createCell(cellIndex++).setCellValue(list.get(i));
			//System.out.println(list.get(i));
		}

		//write this workbook in excel file.
		try {
			FileOutputStream fos = new FileOutputStream("Dummy.xlsx");
			workbook.write(fos);
			System.out.println("cretaed excel ....");
			fos.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}


	}

	public static void main(String[] args) {
		List<String> list = readExcelFile();
		writeExcel(list);

		
	}
}
