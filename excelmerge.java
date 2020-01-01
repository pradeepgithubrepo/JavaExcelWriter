package com.prad.test;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.*;
import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.FileOutputStream;
import java.io.IOException;  

public class excel_mrg {
	// Flags to control writing to Excel
	public static int rowNum = 1;
	public static int headerflag = 0;
	public static int cellcounter = 0;

	public static Sheet sheet = null;
	public static Sheet newSheet = null;
	public static Row row= null;
	public static Cell cellRow = null;
	public static Row Newrow= null;
	public static Cell NewcellRow = null;
	public static Row headerrow=null;
	public static Workbook workbook=null;
	public static CellStyle style=null;
	public static String oldsheetname;
	public static Integer oldrownum;
	public static Integer oldcellnum;
	public static Integer LimitSetter;
	public static XSSFWorkbook Newworkbook;
	public static XSSFWorkbook oldworkbook;
	public static FileInputStream file;

	public static void main(String[] args) throws IOException, InvalidFormatException 
	{
		try
		{
			Newworkbook = new XSSFWorkbook();
			headerStyleSetup();
			headerflag = 1;
			oldworkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/pradeepp/Desktop/javasam/poi-generated-file.xlsx")));
			preProcessor();

			headerflag = 1;
			oldworkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/pradeepp/Desktop/javasam/poi-generated-file1.xlsx")));
			preProcessor();

//			oldworkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/pradeepp/Desktop/javasam/temp.xlsx")));
//			preProcessor();

			FileOutputStream fileOut = new FileOutputStream("/Users/pradeepp/Desktop/javasam/rob_max.xlsx");
			Newworkbook.write(fileOut);
			fileOut.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

	public static void preProcessor() {
		LimitSetter= oldworkbook.getNumberOfSheets();
		for (int i=0 ;i < LimitSetter ; i++)
		{
			sheet = oldworkbook.getSheetAt(i);
			oldsheetname    = sheet.getSheetName();
			newSheet      = Newworkbook.createSheet(oldsheetname);
			newTabProcessor();
		}
	}
	
	public static void headerStyleSetup() {
		style = Newworkbook.createCellStyle();  
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());  
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND); 	
        Font headerFont = Newworkbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		style.setFont(headerFont);
	}
	
	public static void newTabProcessor() {
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) 
		{
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			System.out.println("Old ROW :"+row.getRowNum());
			oldrownum=row.getRowNum();

			while (cellIterator.hasNext()) 
			{
				Cell cell = cellIterator.next();
				System.out.println("old cell-Number : "+cell.getColumnIndex());
				oldcellnum = cell.getColumnIndex();
				Newrow = newSheet.getRow(oldrownum);
				if (Newrow == null) Newrow = newSheet.createRow(oldrownum);
				NewcellRow = Newrow.getCell(oldcellnum);
				if (NewcellRow == null)NewcellRow = Newrow.createCell(oldcellnum);
				System.out.println("Old Cell Value1 :"+cell.getCellType());
				switch (cell.getCellType()){
				case NUMERIC:
					NewcellRow.setCellValue(cell.getNumericCellValue());
					break;
				default :
					NewcellRow.setCellValue(cell.getStringCellValue());
					break;
				}
				if(headerflag == 1) {
					CellStyle newStyle = Newworkbook.createCellStyle();
					newStyle.cloneStyleFrom(cell.getCellStyle());
					NewcellRow.setCellStyle(newStyle);
					headerflag=0;
					
				}

			}
		}
	}

}
