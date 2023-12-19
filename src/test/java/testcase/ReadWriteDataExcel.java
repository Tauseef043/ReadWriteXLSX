package testcase;

import java.util.List;


import org.openqa.selenium.WebDriver;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadWriteDataExcel {

	@Test
	public void RedirectToURL() {

		WebDriver driver = new ChromeDriver();
		System.out.println(System.getProperty("user.dir"));

		System.setProperty("webdriver.chrome.driver",
				System.getProperty("user.dir") + "src\\main\\java\\resources\\chromedriver.exe");
		
//		driver.get("https://google.com");
	}

	@Test(priority=1)
	public void writeData() {
		// Print something on console/screen
		
//		XSSFWorkbook workbook = new XSSFWorkbook();
//		XSSFSheet sheet = workbook.createSheet("Employee Data"); // Replace with desired sheet name
//		// Create a row at a specific index (e.g., 0 for the first row)
//		Row row = sheet.createRow(0);
//
//		// Create cells and set their values
//		Cell cell1 = row.createCell(0);
//		cell1.setCellValue("Data1");
//
//		Cell cell2 = row.createCell(1);
//		cell2.setCellValue("Data2");
//
//		// Continue creating cells as needed
//		try {
//		    FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\java\\resources\\book.xlsx"); // Replace with desired file path
//		    workbook.write(outputStream);
//		    outputStream.close();
//		    System.out.println("Data written to Excel file successfully!");
//		} catch (Exception e) {
//		    e.printStackTrace();
//		}
		try {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        Cell headerCell1 = headerRow.createCell(0);
        headerCell1.setCellValue("id");
        Cell headerCell2 = headerRow.createCell(1);
        headerCell2.setCellValue("name");
        Cell headerCell3 = headerRow.createCell(2);
        headerCell3.setCellValue("salary");
        // Create data rows
        String[][] data = {
                {"1", "tauseef", "1000"},
                {"2", "taha", "1500"},
                {"3", "hamza", "200"},
                {"4", "haider", "400"}
        };
        
        
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i + 1); // Start from row 1 (after header)
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }
        // Write the workbook to a file
        FileOutputStream outputStream =new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\java\\resources\\book.xlsx");
        workbook.write(outputStream);
        outputStream.close();

        System.out.println("Data written to Excel file successfully!");
    } catch (Exception e) {
        e.printStackTrace();
    }

	}
	
	@Test(priority=2)
	public void readData() throws IOException {
		
		FileInputStream inputStream = new FileInputStream(System.getProperty("user.dir") + "\\src\\main\\java\\resources\\book.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		
		XSSFSheet sheet = workbook.getSheetAt(0); // Replace with the index of the sheet you want
		int rows = sheet.getLastRowNum(); // Get the last row number
		for (int i = 0; i <= rows; i++) {
		    Row row = sheet.getRow(i);
		    if (row != null) {
		        int cells = row.getLastCellNum(); // Get the last cell number in the row
		        for (int j = 0; j < cells; j++) {
		            Cell cell = row.getCell(j);
		            if (cell != null) {
		                String cellValue = cell.getStringCellValue(); // Handle data types as needed
		                System.out.print(cellValue + "\t");
		            }
		        }
		        System.out.println(); // Print a newline after each row
		    }
		}
		
		workbook.close();


	}


}
