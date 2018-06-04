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
			int rowcount=0;
			int custid=1;
			do
			{
				customerinfo="{";
				Row nextRow = iterator.next();
				if (rowcount==0)
				{
					nextRow = iterator.next();
					rowcount=rowcount+1;
				}
				else
				{
				custid=(int) custinfo.getRow(rowcount).getCell(0).getNumericCellValue();
				customerinfo= customerinfo+"\"id\": \""+String.valueOf(Math.round(custid))+"\"";
				customerinfo=customerinfo+", "+"\"FirstName\": \""+custinfo.getRow(rowcount).getCell(1).getStringCellValue()+"\"";
				customerinfo=customerinfo+", "+"\"LastName\": \""+custinfo.getRow(rowcount).getCell(2).getStringCellValue()+"\"";
				rowcount=rowcount+1;
				}
		 	 //Adding Address Info in JSON
		 	       Sheet addressinfo=infoworkbook.getSheet("Address");
					Iterator<Row> iteratoraddress=addressinfo.iterator();
					int rowcountaddress=0;
					int addresscount=1;
					do
					{
						
						Row nextRowaddress = iteratoraddress.next();
						if (rowcountaddress==0)
						{
							nextRowaddress = iteratoraddress.next();
						//	rowcountaddress=rowcountaddress+1;
						}
						else
						{
						if (custid==(int) addressinfo.getRow(rowcountaddress).getCell(0).getNumericCellValue())
						{
						customerinfo=customerinfo+", \"Address" + addresscount + "\": \"";
						customerinfo=customerinfo+addressinfo.getRow(rowcountaddress).getCell(1).getStringCellValue()+"\"";
						addresscount=addresscount+1;
						}
						}
						rowcountaddress=rowcountaddress+1;
				 	    
				 	        
					}while (iteratoraddress.hasNext());
			
					//Adding Phone Number
			 	       Sheet phoneinfo=infoworkbook.getSheet("Phone");
						Iterator<Row> iteratorphone=phoneinfo.iterator();
						int rowcountphone=0;
						int phonecount=1;
						do
						{
							
							Row nextRowphone = iteratorphone.next();
							if (rowcountphone==0)
							{
								nextRowphone = iteratorphone.next();
							//	rowcountaddress=rowcountaddress+1;
							}
							else
							{
							if (custid==(int) phoneinfo.getRow(rowcountphone).getCell(0).getNumericCellValue())
							{
							customerinfo=customerinfo+", \"Phone" + phonecount + "\": \"";
							customerinfo=customerinfo+String.valueOf(Math.round(phoneinfo.getRow(rowcountphone).getCell(1).getNumericCellValue()))+"\"}";
							phonecount=phonecount+1;
							}
							}
							rowcountphone=rowcountphone+1;
					 	    
					 	        
						}while (iteratorphone.hasNext());

					System.out.println(customerinfo);
			}while (iterator.hasNext());
			
	}
	

	@Test (priority=3)
	public void postCustomer() throws IOException 
	{
	Response addcustomer=given().
			body(customerinfo).
			when().
			contentType(ContentType.JSON).
			post("http://localhost:3000/customer");		
	}
}