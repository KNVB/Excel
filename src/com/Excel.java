package com;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
public class Excel 
{
	
	public static void main(String[] args) 
	{
		String inputFilePath="G:\\workspace\\test.xlsx";
		String outputFilePath="G:\\workspace\\output.xlsx";
		try 
		{
			 FileInputStream input_document = new FileInputStream(inputFilePath);
			 FileOutputStream output_file =new FileOutputStream(outputFilePath);
             // convert it into a POI object
             XSSFWorkbook workbook = new XSSFWorkbook(input_document); 
             // Read excel sheet that needs to be updated
             XSSFSheet sheet1 = workbook.getSheet("sheet1");
             XSSFRow destRow=sheet1.getRow(3);
 			 XSSFCell cell=destRow.getCell(0);
             
 			 cell.setCellValue(cell.getStringCellValue() + "123");
             //important to close InputStream
             //input_document.close();
             //write changes
             workbook.write(output_file);
             //close the stream
             //output_file.close(); 
             workbook.close();
             			
		} 
		catch (IOException e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}