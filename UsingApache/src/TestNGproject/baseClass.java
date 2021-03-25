package TestNGproject;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;



public class baseClass {
	WebDriver driver;
	
@BeforeSuite
	//login to sathya 
public void setup() {
	System.setProperty("webdriver.gecko.driver","C:\\Users\\Monica\\Selenium\\web driver\\geckodriver.exe");
	 driver = new FirefoxDriver();
	driver.manage().window().maximize();
	String url = "https://www.sathya.in/";
	driver.get(url);
	driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
}

@Test
public void Readvalue() {
	try {
		 FileInputStream fs=new FileInputStream("D:\\Excel\\sathya.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fs);
			Sheet sheet = wb.getSheet("Sheet1");
			Iterator<Row> rowiterator=sheet.iterator();
			while(rowiterator.hasNext()) {
				Row rowvalue=rowiterator.next();
				Iterator<Cell> columniterator=rowvalue.iterator();
				int i=2;
				while(columniterator.hasNext()) {
					if(i%2==0) {
						username.add(columniterator.next().getStringCellValue());
					}
					else {
						password.add(columniterator.next().getStringCellValue());

					}
				
				i++;
				}
			}
	}
	catch(Exception e) {
		System.out.println(e.getMessage());
	}
       

	}
}
