package ApachePOI.Assignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;


public class ExcelDataProcessor {
    @Test
    @Parameters({"namesA","namesB"})
	public static void DataProcessor(String namesA,String namesB ) throws IOException {

    	
    	String[] namesinACol = namesA.split(",");
    	String[] namesinBCol = namesB.split(",");
      Set<String> uniqueNamesA =  new HashSet<String>();
      Set<String> uniqueNamesB =  new HashSet<String>();
      Set<String> duplicateNames =  new HashSet<String>();
      boolean duplicateValue;
      // logic
   
      for (String nameA : namesinACol) {
    	  duplicateValue = false;
    	  for (String nameB: namesinBCol) {

    		   if(nameA.equals(nameB)){

    			   duplicateNames.add(nameA);
    			   duplicateValue = true;

    		   }
    		  
    		   
    	  }
    	  if(!duplicateValue) {
			   uniqueNamesA.add(nameA);
		   }


    }
      
      for(String nameB : namesinBCol) {
    	  duplicateValue = false;
    	  for(String Value : duplicateNames) {
    		  
    		  if(nameB.equals(Value)) {
    			  duplicateValue = true;    			  
    		  }
    		  
    		   
    	  }
    	  if(!duplicateValue) {
    		  uniqueNamesB.add(nameB);
		   }
    	  
      }



         writeExcel(System.getProperty("user.dir")+"/Utils/AssignmentWriteData.xlsx", uniqueNamesA, uniqueNamesB, duplicateNames);
    }




   


    public static void writeExcel(String outputPath, Set<String> uniqueNamesA, Set<String> uniqueNamesB, Set<String> duplicates) throws IOException {


        try {
        	 Workbook workbook = new XSSFWorkbook();
             File file = new File(outputPath);
             FileOutputStream fos = new FileOutputStream(file);
             Sheet sheet = workbook.createSheet("Processed Data");
             int rowIndex = 0;

            Row headerRow = sheet.createRow(rowIndex);
            headerRow.createCell(0).setCellValue("Unique Names-A");
            headerRow.createCell(1).setCellValue("Unique Names-B");
            headerRow.createCell(2).setCellValue("Duplicates Names- A & B");

            
           rowIndex = 1;
            for (String nameA : uniqueNamesA) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(nameA);


            }


            rowIndex = 1;
            for (String nameB : uniqueNamesB) {
                Row row = sheet.getRow(rowIndex++);
                if (row == null) row = sheet.createRow(rowIndex - 1);
                row.createCell(1).setCellValue(nameB);

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

            fos.close();
        } catch ( IOException e) {

        	e.printStackTrace();

             }
    }}

    
    
 
 






