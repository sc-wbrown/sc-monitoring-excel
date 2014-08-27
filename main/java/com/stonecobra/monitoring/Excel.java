package com.stonecobra.monitoring;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * This class allows users to create, read and write Excel Workbooks and Sheets.
 * @author Weslee Brown - wbrown@stonecobra.com
 */
public class Excel {
	public Workbook workbook;
	public Sheet sheet;
	
	/**
	 * This constructor will create an Excel instance to create workbooks and sheets 
	 */
	public Excel() {
		init();
	}
	
	/**
	 * This constructor will create an excel workbook instance from the given excel file
	 * @param name, the name of the excel file
	 */
	public Excel(String filename) {
		init(filename);
	}

	/**
	 * This method will create an instance of the Excel class
	 */
	private void init(){}
	
	/**
	 * This method will create a workbook instance from the given excel file
	 * @param filename, the name of the excel file to create a workbook instance
	 * @return, a workbook instance
	 */
	private Workbook init(String filename) {
		 workbook = openWorkbook(filename);
		 return workbook;
	}
	
	/**
	 * This method will return the user the desired sheet from the workbook
	 * @param workbook, the workbook that contains the sheet
	 * @param page, the sheet of the workbook you wish to make changes to
	 * @return
	 */
	public Sheet getSheet(Workbook workbook, int page) {
		Sheet sheet;
		sheet = workbook.getSheetAt(page);
		return sheet;
	}
	
	/**
	 * This method will create a new sheet in an existing workbook
	 * @param workbook, the workbook that will get the new sheet
	 * @param name, the name of the sheet to create
	 * @return, the sheet instance
	 */
	public Sheet createSheet(Workbook workbook, String name){
		Sheet sheet = workbook.createSheet(name);
		return sheet;
	}
	
	/**
	 * This method will take a sheet and return the rows to iterate
	 * @param sheet, the sheet to obtain the rows from
	 * @return, the row iterator
	 */
	private Iterator<Row> iterateRows(Sheet sheet) {
		Iterator<Row> rowIterator;
		rowIterator = sheet.iterator();
		return rowIterator;
	}
	
	/**
	 * This method will take a row and return the cells to iterate
	 * @param row, the row to obtain the cells from
	 * @return, the cell iterator
	 */
	private Iterator<Cell> iterateCells(Row row) {
		Iterator<Cell> cellIterator;
		cellIterator = row.cellIterator();
		return cellIterator;
	}
	
	/**
	 * This method will iterate through a workbook sheet and print the contents of each cell
	 * @param sheet, the sheet to iterate
	 */
	public void iterateEntireSheet(Sheet sheet){
		Iterator<Row> rowIterator = iterateRows(sheet);
		while( rowIterator.hasNext() ) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = iterateCells(row);
			while( cellIterator.hasNext() ) {
				Cell cell = cellIterator.next();
				
				switch(cell.getCellType()){
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + "\t\t");
					break;
				case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue() + "\t\t");
					break;			
				case Cell.CELL_TYPE_BLANK:
					System.out.print("\t\t");
					break;
				default: System.out.print("\t\t");
					break;	
				}
			}
			System.out.println("");
		}
	}
		
	/**
	 * This method will take an excel file and return a workbook instance 
	 * @param filename, the excel file used to create a workbook instance
	 * @return a workbook instance
	 */
	public Workbook openWorkbook(String filename) {
		Workbook workbook = null;
		FileInputStream file = null;
		try {
			file = new FileInputStream(filename);
			if(filename.toLowerCase().endsWith("xls")) {
				try {
					workbook = new HSSFWorkbook(file);
					file.close();
				}
				catch (IOException e1) {
					e1.printStackTrace();
				}
			}
			else if(filename.toLowerCase().endsWith("xlsx")) {
				try {
					workbook = new XSSFWorkbook(file);
					file.close();
				}
				catch (IOException e2) {
					e2.printStackTrace();
				}	
			}
		} catch (FileNotFoundException e3) {
			e3.printStackTrace();
		}
		return workbook;
	}
	
	/**
	 * This method will iterate through every sheet in the workbook and print out its contents
	 * @param workbook, the workbook 
	 */
	public void readExcel(Workbook workbook) {
		for(int x = 0; x < workbook.getNumberOfSheets(); x++) {
			System.out.println("Sheet " + ++x);
			x--;
			iterateEntireSheet(getSheet(workbook,x));
		}
	}
	
	/**
	 * This method will write the workbook to an output file
	 * @param filename, the name of the file you want to write the workbook to
	 * @param workbook, the workbook you wish to have written to the file
	 */
	public void writeExcelFile(String filename, Workbook workbook) {
		try {
			FileOutputStream out = new FileOutputStream(new File(filename));
			workbook.write(out);
			out.close();
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e2) {
			e2.printStackTrace();
		}
	}
}