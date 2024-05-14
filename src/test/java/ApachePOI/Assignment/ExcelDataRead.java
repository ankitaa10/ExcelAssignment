package ApachePOI.Assignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelDataRead {
	 public static Object[][] getExcelData() throws IOException {
	        // Create a workbook object

	     File file = new File(System.getProperty("user.dir")+"/Utils/AssignmentReadData.xlsx");
         FileInputStream fis = new FileInputStream(file);

    
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);



            int rowCount = sheet.getLastRowNum();
            int columnCount = sheet.getRow(0).getLastCellNum();

	       

	        // Create an object array to store the data
	        Object[][] data = new Object[rowCount][columnCount];

	        // Iterate over the rows and columns
	        for (int i = 0; i < rowCount; i++) {
	            Row row = sheet.getRow(i);
	            for (int j = 0; j < columnCount; j++) {
	                Cell cell = row.getCell(j);
	                data[i][j] = cell.getStringCellValue();
	            }
	        }

	        // Close the workbook
	        workbook.close();

	        return data;
	    }
	 
	 
	 @Test
	 public void testExcelData() throws IOException {
	     // Get the data from the Excel file
	     Object[][] data = ExcelDataRead.getExcelData();
	 }
	}

