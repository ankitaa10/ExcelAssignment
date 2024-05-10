package ApachePOI.Assignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.testng.annotations.Test;

@Test
public class ExcelDataProcessor {
    @Test
	public static void DataProcessor() throws IOException {
		
    
    	String[][] data = readExcel(System.getProperty("user.dir")+"/Utils/AssignmentData.xlsx");
    	
    	String[] names = data[0];
      String[] locations = data[1];
      
      Set<String> uniqueNames = new HashSet<>();
      Set<String> uniqueLocations = new HashSet<>();
      Set<String> duplicates = new HashSet<>();

      
      // clearing duplicates 
      
      for (String name : names) {
        if (!uniqueNames.add(name)) { 
           duplicates.add(name);
           
       }
    }

    
         for (String location : locations) {
        if (!uniqueLocations.add(location)) {
           duplicates.add(location);
        }
        }
      
         writeExcel(System.getProperty("user.dir")+"/Utils/AssignmentWriteData.xlsx", uniqueNames, uniqueLocations, duplicates);
    }
    	
    	
    
	
    
    public static String[][] readExcel(String inputPath) throws IOException {
        
        
        try {
              FileInputStream file = new FileInputStream(new File(inputPath));
        // create new workbook
             Workbook workbook = new XSSFWorkbook(file); 
             //get first sheet 
            Sheet sheet = workbook.getSheetAt(0);
            
            // getting the no of rows and columns in excel sheet
            
           int rowCount = sheet.getLastRowNum();
           int colCount = sheet.getRow(0).getLastCellNum();
           
           
           String[] names = new String[rowCount];
           String[] locations = new String[rowCount];
           
          
          
           
            //iterate through all rows except first row
            for(int i =1; i < (rowCount);i++) {
            	
            	Row row = sheet.getRow(i);
            	
            		names[i] = row.getCell(0).getStringCellValue();
            		locations[i] = row.getCell(1).getStringCellValue();	
            }
            
            
            String[][] excelData = {names,locations};
         
    
            file.close();
         // adding all data to excel data
         return excelData;
        
        
    } catch ( IOException e) {
    	e.printStackTrace();
    }
		return null;
    }
    
    
    public static void writeExcel(String outputPath, Set<String> uniqueNames, Set<String> uniqueLocations, Set<String> duplicates) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(new File(outputPath))) {
            Sheet sheet = workbook.createSheet("Processed Data");
            int rowIndex = 0;
 
            Row headerRow = sheet.createRow(rowIndex);
            headerRow.createCell(0).setCellValue("Unique Names");
            headerRow.createCell(1).setCellValue("Unique Locations");
            headerRow.createCell(2).setCellValue("Duplicates");
 
            //  to get unique names and locations
           rowIndex = 1;
            for (String name : uniqueNames) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(name);
                
                
            }
 
           
            rowIndex = 1;
            for (String location : uniqueLocations) {
                Row row = sheet.getRow(rowIndex++);
                if (row == null) row = sheet.createRow(rowIndex - 1);
                row.createCell(1).setCellValue(location);
               
            }
 
            
            
            // Duplicates
            rowIndex = 1;
            for (String duplicate : duplicates) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) row = sheet.createRow(rowIndex - 1);
                row.createCell(2).setCellValue(duplicate);
                rowIndex++;
            }
 
            workbook.write(fos);
        }
    }
    
    
    	
    	
 	
    	
    }
 
   

     
 


