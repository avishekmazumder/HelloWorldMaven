package com.avi.excelcomparator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelComparator {
	public static int resultSheetCurrentRow=0;
	public static List<String> diffString;
	public static void main(String args[]) {
		
		try {
			FileInputStream excellFile1 = new FileInputStream(new File(
                    "E:\\ExcelComparison\\Enrichment.xlsx"));
            FileInputStream excellFile2 = new FileInputStream(new File(
                    "E:\\ExcelComparison\\EnrichmentModified.xlsx"));
            
            XSSFWorkbook workbook1 = new XSSFWorkbook(excellFile1);
            XSSFWorkbook workbook2 = new XSSFWorkbook(excellFile2);
            
         // Get first/desired sheet from the workbook
            XSSFSheet sheet1 = workbook1.getSheet("First");
            XSSFSheet sheet2 = workbook2.getSheet("First");
            
            XSSFWorkbook result = new XSSFWorkbook();
            XSSFSheet resultSheet1 = result.createSheet("First");
            
            compareAndUpdateDiff(sheet1,sheet2,resultSheet1);
            int lastCellNum = resultSheet1.getRow(0).getLastCellNum();
    		for(int j=0;j<lastCellNum;j++) {
    			resultSheet1.setColumnWidth(j, 10000);
			}
            
            FileOutputStream fileOut = new FileOutputStream("E:\\ExcelComparison\\poi-generated-file.xlsx");
            result.write(fileOut);
            fileOut.close();
            
		} catch(Exception e) {
			
		} finally {
		}
	}
	
	public static void compareAndUpdateDiff(XSSFSheet sheet1, XSSFSheet sheet2, XSSFSheet resultSheet) {
		int sheet1FirstRow = sheet1.getFirstRowNum();
		int sheet2LastRow = sheet1.getLastRowNum();
		
		XSSFRow headerRow = sheet1.getRow(sheet1FirstRow);
		
		populateFirstRow(headerRow,resultSheet);
		
		for (int i=sheet1FirstRow+1;i<sheet2LastRow;i++) {
			XSSFRow row1=sheet1.getRow(i);
			XSSFRow row2=sheet2.getRow(i);
			
			boolean flag = compareTwoRows(row1,row2);
			
			if(flag) {
				XSSFRow newRow = resultSheet.createRow(resultSheetCurrentRow);
				resultSheetCurrentRow++;
				for(int k=0;k<diffString.size();k++) {
					XSSFCell newCell = newRow.createCell(k);
					newCell.setCellValue(diffString.get(k));
				}
			}
			
		}
	}
	
	private static boolean compareTwoRows(XSSFRow row1, XSSFRow row2) {
		int firstCell1 = row1.getFirstCellNum();
        int lastCell1 = row1.getLastCellNum();
        diffString = new ArrayList<>();
        boolean flag=false;
        
        for(int i=firstCell1;i<lastCell1;i++) {
        	 XSSFCell cell1 = row1.getCell(i);
             XSSFCell cell2 = row2.getCell(i);
             
            String rslt = (String) compareTwoCells(cell1,cell2);
            
            if(!rslt.equals(cell1.getStringCellValue())) {
            	flag=true;
            	diffString.add(rslt);
            } else {
            	if(flag) {
            		diffString.add("");
            	}
            }
        }
        return flag;
		
	}

	private static String compareTwoCells(XSSFCell cell1, XSSFCell cell2) {
		String val1=cell1.getStringCellValue();
		String val2=cell2.getStringCellValue();
		
		if(val1.equals(val2)) {
			return val1;
		} else {
			return "First File content::::[["+val1+"]]\nSecond File Content::::[["+val2+"]]";
		}		
	}

	public static void populateFirstRow(XSSFRow headerRow,XSSFSheet resultSheet) {
		
		XSSFRow firstRow=resultSheet.createRow(resultSheetCurrentRow);
		resultSheetCurrentRow++;
		int i=0;
		for (Cell cell : headerRow) {
			Cell resCell = firstRow.createCell(i++);
			resCell.setCellValue(cell.getStringCellValue());
			
		}
		
	}

}
