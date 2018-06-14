package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel1 {

	public static void main(String[] args) 
	{
		/*
		String inputFilePath="G:\\workspace\\test.xlsx";
		String outputFilePath="G:\\workspace\\output.xlsx";
		*/
		
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
			XSSFRow destRow=sheet1.getRow(3);
			sheet1.shiftRows(3, 3, 1);
			
			XSSFRow sourceRow=sheet2.getRow(12);
			XSSFCell cell=destRow.getCell(0);
			
			List<XSSFRow> sourceRows=new ArrayList<XSSFRow>();
 			CellCopyPolicy cellCopyPolicy=new CellCopyPolicy();
 			cellCopyPolicy.setCopyCellFormula(true);
 			cellCopyPolicy.setCopyCellStyle(true);
 			cellCopyPolicy.setCopyCellValue(true);
 			sourceRows.add(sourceRow);
 			
 			sheet1.copyRows(sourceRows, 3,cellCopyPolicy);
             
 			workbook.write(new FileOutputStream(outputFile));
			workbook.close();
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}
	}

}
