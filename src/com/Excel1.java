package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import common.CalendarUtility;
import common.MonthlyCalendar;
import common.MyCalendar;

public class Excel1 {

	public static void main(String[] args) 
	{
		/*
		String inputFilePath="G:\\workspace\\template.xlsx";
		String outputFilePath="G:\\workspace\\output.xlsx";
		*/
		String inputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\template.xlsx";
		String outputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\output.xlsx";
		
		File inputFile=new File(inputFilePath);
		File outputFile=new File(outputFilePath);
		String monthName;
		
		try 
		{
			Files.copy(inputFile.toPath(), outputFile.toPath(),StandardCopyOption.REPLACE_EXISTING);
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(outputFile));
			XSSFSheet sheet1 = workbook.getSheet("sheet1");
			XSSFSheet sheet2 = workbook.getSheet("sheet2");
		    XSSFSheetConditionalFormatting sheet1cf = sheet1.getSheetConditionalFormatting(); 
            XSSFSheetConditionalFormatting sheet2cf = sheet2.getSheetConditionalFormatting(); 
 			
            
            XSSFCell cell=sheet1.getRow(1).getCell(1);
 			Calendar calendar = new GregorianCalendar(); 			
 			monthName=sheet2.getRow(calendar.get(Calendar.MONTH)).getCell(0).getStringCellValue()+" "+calendar.get(Calendar.YEAR);
 			cell.setCellValue(monthName);
 			
 		    CalendarUtility calendarUtility=new CalendarUtility();
            MonthlyCalendar mc=calendarUtility.getMonthlyCalendar(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH));
            for (int i=1;i<=mc.length;i++)
    		{
            	MyCalendar myCalendar=mc.getMonthlyCalendar().get(i);
            	cell=sheet1.getRow(3).getCell(i);
            	if (myCalendar.isPublicHoliday())
            		cell.setCellValue("PH");
            	cell=sheet1.getRow(4).getCell(i);
            	"Su"
            	"M"
            	"T"
            	"W"
                "Th"
            	"F"
            	"S"    	
            	cell=sheet1.getRow(5).getCell(i);
            	cell.setCellValue(i);
    		} 			
 			
 			
 			sheet1.shiftRows(6, 6, 1);
 			
 			XSSFRow sourceRow=sheet2.getRow(12);
			List<XSSFRow> sourceRows=new ArrayList<XSSFRow>();
 			CellCopyPolicy cellCopyPolicy=new CellCopyPolicy();
 			/*cellCopyPolicy.setCopyCellFormula(true);
 			cellCopyPolicy.setCopyCellStyle(true);
 			cellCopyPolicy.setCopyCellValue(true);*/
 			sourceRows.add(sourceRow);
 			
 			sheet1.copyRows(sourceRows, 6,cellCopyPolicy);
 			
 			cell=sheet1.getRow(6).getCell(1);
 			cell.setCellValue("a");
            for (int idx = 0; idx < sheet2cf.getNumConditionalFormattings(); idx++) 
            { 
            	XSSFConditionalFormatting cf = sheet2cf.getConditionalFormattingAt(idx);
            	CellRangeAddress[]ranges=cf.getFormattingRanges();
            	System.out.println("Before Range="+cf.getFormattingRanges().length);
            	ranges[0].setFirstRow(6);
            	ranges[0].setLastRow(6);
            	cf.setFormattingRanges(ranges);
            	
/*            	System.out.println("Range="+cf.getFormattingRanges().length);
            	System.out.println("Range="+cf.getFormattingRanges()[0]);*/
            	sheet1cf.addConditionalFormatting(cf);

/*            	System.out.println("idx="+idx);
            	System.out.println("length="+cf.getFormattingRanges().length);
            	System.out.println("Range="+cf.getFormattingRanges()[0].getFirstColumn()+":"+cf.getFormattingRanges()[0].getFirstRow()+"-"+cf.getFormattingRanges()[0].getLastColumn()+":"+cf.getFormattingRanges()[0].getLastRow());
            	
            	for (int j=0;j<cf.getNumberOfRules();j++)
            	{
            		XSSFConditionalFormattingRule rule=cf.getRule(j);
            		System.out.println("j="+j);
            		System.out.println("formula1="+rule.getFormula1());
            		System.out.println("formula2="+rule.getFormula2());
            		System.out.println("ConditionType="+rule.getConditionType());
            		System.out.println("ConditionFilterType="+rule.getConditionFilterType());
            		System.out.println("ColorScaleType="+rule.getColorScaleFormatting());
            		System.out.println("=========================================");
            	}*/
            } 
        
 			workbook.write(new FileOutputStream(outputFile));
			workbook.close();
			//System.out.println(monthName);
			System.out.println("Done");
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}
	}

}
