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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel1 {

	public static void main(String[] args) 
	{
		String inputFilePath="G:\\workspace\\template.xlsx";
		String outputFilePath="G:\\workspace\\output.xlsx";
		
		
		/*
		String inputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\template.xlsx";
		String outputFilePath="D:\\Inetpub\\wwwroot\\roster-sample\\output.xlsx";
		*/
		File inputFile=new File(inputFilePath);
		File outputFile=new File(outputFilePath);
		String monthName;
		try 
		{
			Files.copy(inputFile.toPath(), outputFile.toPath(),StandardCopyOption.REPLACE_EXISTING);
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(outputFile));
			XSSFSheet sheet1 = workbook.getSheet("sheet1");
			XSSFSheet sheet2 = workbook.getSheet("sheet2");
 			Calendar calendar = new GregorianCalendar();
 			XSSFCell cell=sheet1.getRow(1).getCell(1);
 			monthName=sheet2.getRow(calendar.get(Calendar.MONTH)).getCell(0).getStringCellValue()+" "+calendar.get(Calendar.YEAR);
 			cell.setCellValue(monthName);
 			sheet1.shiftRows(6, 6, 1);
 			XSSFRow destRow=sheet1.getRow(6);
 			XSSFRow sourceRow=sheet2.getRow(12);

 			
			List<XSSFRow> sourceRows=new ArrayList<XSSFRow>();
 			CellCopyPolicy cellCopyPolicy=new CellCopyPolicy();
 			/*cellCopyPolicy.setCopyCellFormula(true);
 			cellCopyPolicy.setCopyCellStyle(true);
 			cellCopyPolicy.setCopyCellValue(true);*/
 			sourceRows.add(sourceRow);
 			
 			sheet1.copyRows(sourceRows, 6,cellCopyPolicy);

 			
 			workbook.write(new FileOutputStream(outputFile));
			workbook.close();
			System.out.println(monthName);
			System.out.println("Done");
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}
	}

}
