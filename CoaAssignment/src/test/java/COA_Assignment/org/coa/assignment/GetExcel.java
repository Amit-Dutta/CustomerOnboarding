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

	@Test
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

}
