package com.prad.test;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.*;
import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.CellStyle;  
import org.apache.poi.ss.usermodel.FillPatternType;  
import org.apache.poi.ss.usermodel.IndexedColors;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel {
	private static String[] columns = {"Name","Name1"};

	// Flags to control writing to Excel
	public static int rowNum = 1;
	public static int cellcounter = 0;
	public static int headerflag=0;
	public static int newTabFlag=0;
	public static int sheetNum=0;
	public static int boundrySetterLower=0;
	public static int boundrySetterUpper=0;

	public static String item;
	public static String itemIndex;
	public static String itemCode;
	public static String writeMode;
	public static Sheet sheet = null;
	public static Sheet sheet1 = null;
	public static Sheet sheet2= null;
	public static Sheet sheet3= null;
	public static Sheet sheet4= null;
	public static Sheet indexSheet = null;
	public static Row row= null;
	public static Cell cellRow = null;
	public static Row headerrow=null;
	public static Iterator iterator=null;
	public static Iterator indexIterator=null;
	public static Workbook workbook=null;
	public static CellStyle style=null;

	public static void main(String[] args) throws IOException, InvalidFormatException {
		// Create a Workbook
		workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		headerStyleSetup();
//		Create a Index Page
		createIndexTab();
		
		ArrayList<String> mlist = new ArrayList<String>();
		mlist.add("$a$tableschema");
		mlist.add("@create table");
		mlist.add("col1 ,String");
		mlist.add("cols2 , int");
		mlist.add("$b$distinctval");
		mlist.add("@algorithm");
		mlist.add("abc");
		mlist.add("xyz");
		mlist.add("@kappan");
		mlist.add("123");
		mlist.add("456");
		mlist.add("778");
		mlist.add("@Taj");
		mlist.add("HSSFWorkbook");
		mlist.add("individual");
		mlist.add("cells");
		mlist.add("XSSFFont");
		mlist.add("BOLDWEIGHT_BOLD");
		mlist.add("$c$nullblank");
		mlist.add("@ndc");
		mlist.add("899");
		mlist.add("23%");
		mlist.add("43%");
		mlist.add("@Apache Poi");
		mlist.add("Learn to read excel, write excel, ");
		mlist.add("trusted library among many other ope");
		mlist.add("read and write MS Word a");
		mlist.add("$d$horizontal");
		mlist.add("#Name");
		mlist.add("aage");
		mlist.add("rroll");
		mlist.add("#Pradeep");
		mlist.add("29");
		mlist.add("947");
		mlist.add("#Gokul");
		mlist.add("45");
		mlist.add("777");

		iterator = mlist.iterator();

		while(iterator.hasNext()) {

			item = 	(String) iterator.next();
			itemIndex = item.substring(0, 1);

			if (itemIndex.equals("$"))
			{
				itemCode = item.substring(0, 3);
				System.out.println("itemCode"+itemCode);
				if(itemCode.equals("$a$"))
				{
					sheet1 = workbook.createSheet(item.substring(3));
					sheetNum=1;
					updateItertorFlags();
				}
				if(itemCode.equals("$b$"))
				{
					setUpperboundryapplyautosize();
					sheet2 = workbook.createSheet(item.substring(3));
					sheetNum=2;
					updateItertorFlags();
					setboundrySetterLower();
				}
				if(itemCode.equals("$c$"))
				{
					setUpperboundryapplyautosize();
					sheet3 = workbook.createSheet(item.substring(3));
					sheetNum=3;
					updateItertorFlags();
				}

				if(itemCode.equals("$d$"))
				{
					setUpperboundryapplyautosize();
					sheet4 = workbook.createSheet(item.substring(3));
					sheetNum=4;
					updateItertorFlags();
				}
			}
			
			switch (itemIndex) {
			case "@":writeMode="v";break;
			case "#":writeMode="h";break;}

			if ((itemIndex.equals("@") || itemIndex.equals("#")) && newTabFlag != 1)
			{
				headerflag=1;
				switch (writeMode) {
				case "v":
					rowNum = 0;
					cellcounter++;
					break;
				case "h":
					rowNum ++;
					cellcounter = 0;
					break;
				}
			}

			if (sheetNum==1)
			{
				sheet = sheet1;
				switch (writeMode) {
				case "v":
					System.out.println("coming inside sheetNum,item1"+item);
					processorforv();
					break;
				case "h":
					System.out.println("case h in sheet1");
					processorforh();
					break;
				}
			}

			if (sheetNum==2)
			{
				sheet = sheet2;
				switch (writeMode) {
				case "v":
					System.out.println("coming inside sheetNum,item2"+item);
					processorforv();
					break;
				case "h":
					System.out.println("case h in sheet2");
					processorforh();
					break;
				}
			}

			if (sheetNum==3)
			{
				sheet = sheet3;
				switch (writeMode) {
				case "v":
					System.out.println("coming inside sheetNum,item3"+item);
					processorforv();
					break;
				case "h":
					System.out.println("case h in sheet3");
					processorforh();
					break;
				}
			}

			if (sheetNum==4)
			{
				sheet = sheet4;
				switch (writeMode) {
				case "v":
					processorforv();
					break;
				case "h":
					System.out.println("case h in sheet4");
					processorforh();
					break;
				}
			}

			if (headerflag==1)
			{
				System.out.println("cellcounter in header:"+cellcounter);
				headerrow.getCell(cellcounter).setCellStyle(style);
				if(writeMode.equals("v"))headerflag=0;
			}
		}
		
		//Set Auto Size for the last Tab
		setUpperboundryapplyautosize();

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("/Users/pradeepp/Desktop/javasam/poi-generated-file.xlsx");
		workbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		workbook.close();
	}


	public static void updateItertorFlags(){
		rowNum = 0;
		cellcounter = 0;
		item = 	(String) iterator.next();
		itemIndex = item.substring(0, 1);
		headerflag=1;
		newTabFlag=1;
		setboundrySetterLower();
	}

	public static void setnewTabValuetoZero(){newTabFlag=0;}

	public static void writeToCell(){
		cellRow = row.getCell(cellcounter);
		if (cellRow == null)cellRow = row.createCell(cellcounter);
		cellRow.setCellValue(item);}
	
	public static void processorforv() {
		setnewTabValuetoZero();
		row = sheet.getRow(rowNum);
		if (row == null) row = sheet.createRow(rowNum++);else {rowNum++;}
		if (headerflag==1) {headerrow=sheet.getRow(0);item=item.substring(1);}
		writeToCell();
	}
	
	public static void processorforh() {
		setnewTabValuetoZero();
		row = sheet.getRow(rowNum);
		if (row == null) row = sheet.createRow(rowNum);else {cellcounter++;}
		if (headerflag==1) {headerrow=sheet.getRow(0);item=item.substring(1);}
		writeToCell();
	}
	
	public static void createIndexTab() {
		ArrayList<String> llist = new ArrayList<String>();
		llist.add("!partition");
		llist.add("value");
		llist.add("!File");
		llist.add("abc.txt");
		llist.add("!Total Rec");
		llist.add("353656");
		
		indexSheet = workbook.createSheet("Index");
		rowNum = 9;
		cellcounter = 5;
		setboundrySetterLower();
		createIndexTabCreateRow();
		createIndexTabSetter("Description");
		applyStyle();
		cellcounter++;
		createIndexTabSetter("Info");
		applyStyle();
		indexIterator = llist.iterator();

		while(indexIterator.hasNext()) {
			item = 	(String) indexIterator.next();
			itemIndex = item.substring(0, 1);
			
			if (itemIndex.equals("!")) {
				rowNum ++;
				cellcounter = 5;
				item=item.substring(1);
			}
			createIndexTabCreateRow();
			createIndexTabSetter(item);
			cellcounter++;
		}
		sheet=indexSheet;
		setUpperboundryapplyautosize();
	}
	
	public static void createIndexTabSetter(String msg) {
		item=msg;
		writeToCell();
	}
	
	public static void createIndexTabCreateRow() {
		row = indexSheet.getRow(rowNum);
		if (row == null) row = indexSheet.createRow(rowNum);
	}
	
	public static void headerStyleSetup() {
		style = workbook.createCellStyle();  
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());  
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND); 	
        Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		style.setFont(headerFont);
	}
	
	public static void applyStyle() {cellRow.setCellStyle(style);}
	public static void setboundrySetterLower() {boundrySetterLower=cellcounter;}
	public static void setUpperboundryapplyautosize() {boundrySetterUpper=cellcounter;applyAutoSize();}
	
	public static void applyAutoSize() {
		for(int i = boundrySetterLower; i <= boundrySetterUpper; i++) {sheet.autoSizeColumn(i);}}
	
}
