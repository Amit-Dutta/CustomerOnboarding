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
			Workbook infoworkbook = new XSSFWorkbook(inputstream);
			Sheet custinfo=infoworkbook.getSheet("CustomerInfo");
			Iterator<Row> iterator=custinfo.iterator();
			int rowcount=1;
			int custid=1;
			while (iterator.hasNext())
			{
				customerinfo="";
				Row nextRow = iterator.next();
				if (rowcount==1)
				{
					nextRow = iterator.next();
					rowcount=rowcount+1;
				}
				else
				{
				custid=(int) custinfo.getRow(rowcount).getCell(0).getNumericCellValue();
				customerinfo="\"id\": \"";
				customerinfo= customerinfo+String.valueOf(Math.round(custid))+"\"";
				customerinfo=customerinfo+", "+"\"FirstName\": \""+custinfo.getRow(rowcount).getCell(1).getStringCellValue()+"\"";
				customerinfo=customerinfo+", "+"\"LastName\": \""+custinfo.getRow(rowcount).getCell(2).getStringCellValue()+"\"";
				rowcount=rowcount+1;
				}
		 	        System.out.println(customerinfo);
			}
			
	}
}