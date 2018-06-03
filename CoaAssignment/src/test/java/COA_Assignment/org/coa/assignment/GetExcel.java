package COA_Assignment.org.coa.assignment;

import com.jayway.restassured.http.ContentType;
import com.jayway.restassured.response.Response;
import static com.jayway.restassured.RestAssured.*;

import java.util.Scanner;

import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class GetExcel {

	public String excellocation;
	public String customerinfo;

	@Test (priority=1)
public void getCustomerInfo() {
		
		Scanner scan = new Scanner (System.in);
		System.out.print("Enter Client ID: ");  
		String stat = scan.next();   
	
		excellocation = given().
				when().
				get("http://localhost:3000/posts").
				then().
				contentType(ContentType.JSON).
				extract().
				path("[" + stat + "].location");

		System.out.println("Excel location: "+excellocation);
}

	@Test (priority=2)
	public void parseExcel() throws IOException 
	{
			String excelFilePath = excellocation;
			FileInputStream inputstream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = new XSSFWorkbook(inputstream);
			Sheet custinfo=workbook.getSheet("CustomerInfo");
			Iterator<Row> iterator=custinfo.iterator();
			
			while (iterator.hasNext())
			{
				Row nextRow = iterator.next();
				customerinfo="Value: ";
		        Iterator<Cell> cellIterator = nextRow.cellIterator();
		 	        while (cellIterator.hasNext()) 
		 	{
		 	        	Cell cell=cellIterator.next();
		 	        	
		 	        	switch (cell.getCellType()) 
		 	        	{
	                    case Cell.CELL_TYPE_STRING:
	                    customerinfo= customerinfo+cell.getStringCellValue();
	                    break;
	                    case Cell.CELL_TYPE_BOOLEAN:
	                    	customerinfo= customerinfo+String.valueOf(cell.getStringCellValue());
	                        break;
	                    case Cell.CELL_TYPE_NUMERIC:
	                    	customerinfo= customerinfo + String.valueOf(cell.getNumericCellValue());
	                        break;
	                }
			}
		 	        System.out.println(customerinfo);
			}
			
			
	}

	
}
