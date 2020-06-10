package com.excelReading;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_Data {

	
	
	public static void main(String ar[]) throws EncryptedDocumentException, IOException {
		
		
		File file=new File("E:/eBookData.xlsx");
		
		FileInputStream fis=new FileInputStream(file);
		
		XSSFWorkbook workBook=(XSSFWorkbook) WorkbookFactory.create(fis);
		
		XSSFSheet getSheet=workBook.getSheetAt(0);
		
		System.out.println("Last row num : " + getSheet.getLastRowNum());
		
		System.out.println("First row num : " + getSheet.getFirstRowNum());
		
		int rowCount=getSheet.getLastRowNum() - getSheet.getFirstRowNum() + 1;
		
		System.out.println(rowCount);
		
		
		for(int i=0;i<=rowCount;i++) {
			
			XSSFRow row=getSheet.getRow(i);
			
			
			for(int j=0;j<=row.getLastCellNum();j++) {
				
				XSSFCell cell=row.getCell(j);
				
				
				System.out.print(cell  +  "      \t       ");
				
			}
			
			System.out.println();
			
		}
		
		workBook.close();
		fis.close();
		
		
		
	}
}
