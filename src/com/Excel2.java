package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2 {

	public static void main(String[] args) 
	{
		/*String inputFilePath="G:\\workspace\\template.xlsx";
		String outputFilePath="G:\\workspace\\output.xlsx";*/
		
		String inputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\template.xlsx";
		String outputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\output.xlsx";
		
		File inputFile=new File(inputFilePath);
		File outputFile=new File(outputFilePath);
		
		try
		{
			Files.copy(inputFile.toPath(), outputFile.toPath(),StandardCopyOption.REPLACE_EXISTING);
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(outputFile));
			XSSFSheet sheet1 = workbook.getSheet("sheet1");
			XSSFSheet sheet2 = workbook.getSheet("sheet2");
			XSSFSheetConditionalFormatting sheet1cf = sheet1.getSheetConditionalFormatting(); 
	        XSSFSheetConditionalFormatting sheet2cf = sheet2.getSheetConditionalFormatting(); 
			
			XSSFCell cell;
			XSSFRow sourceRow=sheet2.getRow(12);
			List<XSSFRow> sourceRows=new ArrayList<XSSFRow>();
 			CellCopyPolicy cellCopyPolicy=new CellCopyPolicy();
 			sourceRows.add(sourceRow);
 			int startRowNum=6,i;
 			
 			 
 			/*for (i=0;i<3;i++)
 			{
 				sheet1.shiftRows(startRowNum+i, sheet1.getLastRowNum(), 1);
 				sheet1.copyRows(sourceRows, startRowNum+i,cellCopyPolicy);
 				cell=sheet1.getRow(startRowNum+i).getCell(1);
 				cell.setCellValue("a");
 				
 				
			}*/
 			
 			for (int idx = 0; idx <sheet2cf.getNumConditionalFormattings(); idx++) 
            { 
 				XSSFConditionalFormatting cf = sheet2cf.getConditionalFormattingAt(idx);
 				CellRangeAddress[]ranges=cf.getFormattingRanges();
 				CellRangeAddress[]r2=new CellRangeAddress[ranges.length];
 				
 				for (int j=0;j<ranges.length;j++)
 				{
 					r2[j]=ranges[j].copy();
 				}
            	r2[0].setFirstRow(startRowNum);
            	r2[0].setLastRow(startRowNum+2);
            	cf.setFormattingRanges(r2);
            	sheet1cf.addConditionalFormatting(cf);
            }
			workbook.write(new FileOutputStream(outputFile));
			workbook.close();
		}
		catch (IOException e) 
		{
			e.printStackTrace();
		}
		
	}

}
