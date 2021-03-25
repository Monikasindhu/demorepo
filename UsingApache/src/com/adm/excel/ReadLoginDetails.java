package com.adm.excel;
//Author Name : Monika R
/*Using Apache Poi jar to read value from Excel sheet.
 *Using Firefox Browser to Login and Logout from the Flipkart website.
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;

public class ReadLoginDetails {
	static List<String> username= new ArrayList<String>();
	static List<String> password= new ArrayList<String>();
	//To read value from Excel
	public void Readvalue() {
	try {
		 FileInputStream fs=new FileInputStream("D:\\Excel\\Rate1.xlsx");
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
	//Login to the flipkart website using webdriver
	public void login(String un,String pw) {
		System.setProperty("webdriver.gecko.driver","C:\\Users\\Monica\\Selenium\\web driver\\geckodriver.exe");

		 WebDriver driver = new FirefoxDriver();
	        driver.manage().window().maximize();
	        String url = "https://www.flipkart.com/";
	        driver.get(url);
	        driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);  
	        Actions builder =new Actions(driver);
	        WebElement Uname= driver.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div[2]/div/form/div[1]/input"));
	        Uname.sendKeys(un);
	        WebElement pword = driver.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div[2]/div/form/div[2]/input"));
	        pword.sendKeys(pw);
	      
	        WebElement loginFlipkart= driver.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div[2]/div/form/div[3]/button"));
	        builder.click(loginFlipkart).build().perform();
	        driver.quit();

	}
	public void executeTest() {
		for(int i=0;i<username.size();i++) {
			login(username.get(i),password.get(i));
		}
	}
	

	public static void main(String[] args) {
		ReadLoginDetails rd=new ReadLoginDetails();
		rd.Readvalue();
      System.out.println("username of flipkart: "+username);
      System.out.println("password of flipkart: "+password);
      rd.executeTest();
			}
}
