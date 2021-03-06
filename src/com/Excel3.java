package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFColorScaleFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFFontFormatting;
import org.apache.poi.xssf.usermodel.XSSFPatternFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel3 
{
	public static void main(String[] args) 
	{

		String outputFilePath="G:\\workspace\\output.xlsx";
		//String outputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\output.xlsx";
		File outputFile=new File(outputFilePath);
		try 
		{
			XSSFColor bgColor;
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(outputFile));
			XSSFSheet sheet1 = workbook.getSheet("sheet1");
			
			XSSFSheetConditionalFormatting sheet1cf = sheet1.getSheetConditionalFormatting(); 
			XSSFPatternFormatting fillPattern;
			CellRangeAddress[] sheet1Range = {CellRangeAddress.valueOf("b7:Af12")};
 
			XSSFConditionalFormattingRule aShiftRule = sheet1cf.createConditionalFormattingRule(ComparisonOperator.EQUAL,"\"a\"");
			XSSFConditionalFormattingRule bxShiftRule = sheet1cf.createConditionalFormattingRule("SEARCH(\"b\",b7:af11)");
			XSSFConditionalFormattingRule cShiftRule = sheet1cf.createConditionalFormattingRule(ComparisonOperator.EQUAL,"\"c\"");
			XSSFConditionalFormattingRule dxShiftRule = sheet1cf.createConditionalFormattingRule("SEARCH(\"d\",b7:af11)");
			XSSFConditionalFormattingRule oShiftRule = sheet1cf.createConditionalFormattingRule(ComparisonOperator.EQUAL,"\"O\"");
			fillPattern = aShiftRule.createPatternFormatting();
			
			bgColor = new XSSFColor(new java.awt.Color(255, 153, 204));
			fillPattern.setFillBackgroundColor(bgColor);
			sheet1cf.addConditionalFormatting(sheet1Range, aShiftRule);

			fillPattern = bxShiftRule.createPatternFormatting();
			bgColor = new XSSFColor(new java.awt.Color(255, 255, 204));
			fillPattern.setFillBackgroundColor(bgColor);
			sheet1cf.addConditionalFormatting(sheet1Range, bxShiftRule);
			
			bgColor = new XSSFColor(new java.awt.Color(204,255,204)); 
			fillPattern = cShiftRule.createPatternFormatting();
			fillPattern.setFillBackgroundColor(bgColor);
			sheet1cf.addConditionalFormatting(sheet1Range, cShiftRule);
			
			bgColor = new XSSFColor(new java.awt.Color(204,255,255));
			fillPattern = dxShiftRule.createPatternFormatting();
			fillPattern.setFillBackgroundColor(bgColor);
			sheet1cf.addConditionalFormatting(sheet1Range, dxShiftRule);
			
			bgColor = new XSSFColor(new java.awt.Color(255,255,255));
			fillPattern =oShiftRule.createPatternFormatting();
			fillPattern.setFillBackgroundColor(bgColor);
			sheet1cf.addConditionalFormatting(sheet1Range, oShiftRule);
			workbook.write(new FileOutputStream(outputFile));
			workbook.close();
		}
		catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}
	}
}
