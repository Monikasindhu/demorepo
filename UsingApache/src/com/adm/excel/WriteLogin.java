
package com.adm.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class WriteLogin {

	
 
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File file=new File("D:\\Excel\\write.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sh=wb.getSheetAt(0);
		

		sh.createRow(0).createCell(0).setCellValue("Product Description");
		sh.getRow(0).createCell(1).setCellValue("Sales Package");
		sh.getRow(0).createCell(2).setCellValue("Model Number");
		sh.getRow(0).createCell(3).setCellValue("Model Name");
		sh.getRow(0).createCell(4).setCellValue("Dial Color");
		sh.getRow(0).createCell(5).setCellValue("Water Resistant");
		
		//Login to flipkart website
		System.setProperty("webdriver.gecko.driver","C:\\Users\\Monica\\Selenium\\web driver\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
        driver.manage().window().maximize();
        String url = "https://www.flipkart.com/";
        driver.get(url);
        driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);  
        
        WebElement Uname= driver.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div[2]/div/form/div[1]/input"));
        Uname.sendKeys("monikasindhu45@gmail.com");
        WebElement pword = driver.findElement(By.xpath("//input[@class='_2IX_2- _3mctLh VJZDxU']"));
        pword.sendKeys("monika@06");
        WebElement loginbutton= driver.findElement(By.xpath("//button[@class='_2KpZ6l _2HKlqd _3AWRsL']"));
        loginbutton.click();
        
        //driver.navigate().refresh();
        //Navigate to smart watches
        WebDriverWait wait = new WebDriverWait(driver, 30);
        Actions builder =new Actions(driver);
        WebElement women = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div/div/div[2]/div/ul/li[4]/span")));
		builder.moveToElement(women).build().perform();
		//driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		WebElement watches =wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div/div/div[2]/div/ul/li[4]/ul/li/ul/li[3]/ul/li[13]/a")));
		watches.click();
		String old_window=driver.getWindowHandle();
        String Before_Xpath = "(//div[@class='_3wUdata53n'])[";
		String After_Xpath="]";
		//		Opening 25 watches products and writing 5 product description in the Excel
		for(int i =1; i<=24; i++)
		{
			String xpath=Before_Xpath+i+After_Xpath;
			WebElement all_watches=wait.until(ExpectedConditions.elementToBeClickable(By.xpath(xpath)));
			all_watches.click();
			Set<String> allwindows=driver.getWindowHandles();
			for (String newwindow : allwindows) {
				driver.switchTo().window(newwindow);
			}
			

			XSSFRow row=sh.createRow(i);
			
			WebElement read_more = driver.findElement(By.xpath("//button[@class='_2AkmmA uSQV49']"));
			read_more.click();
			
			
			
			WebElement Product_Description=driver.findElement(By.xpath("//span[@class='_35KyD6']"));
			String Product_Description_value=Product_Description.getText();
			sh.getRow(i).createCell(0).setCellValue(Product_Description_value);


			WebElement SalesPackage=driver.findElement(By.xpath("//td[@class='_3-wDH3 col col-3-12' and contains (text(), 'Sales Package')]//following::td[1]"));
			String SalesPackage_value=SalesPackage.getText();
			sh.getRow(i).createCell(1).setCellValue(SalesPackage_value);
						
						
			WebElement ModelNumber=driver.findElement(By.xpath("//td[@class='_3-wDH3 col col-3-12'][contains (text(), 'Model Number')]//following::td[1]"));
			String ModelNumber_value=ModelNumber.getText();
			sh.getRow(i).createCell(2).setCellValue(ModelNumber_value);
			
			
			WebElement ModelName=driver.findElement(By.xpath("//td[@class='_3-wDH3 col col-3-12'][contains (text(), 'Model Name')]//following::td[1]"));
			String ModelName_value=ModelName.getText();
			sh.getRow(i).createCell(3).setCellValue(ModelName_value);
			
			
			WebElement DialShape=driver.findElement(By.xpath("//td[@class='_3-wDH3 col col-3-12'][contains (text(), 'Dial Shape')]//following::td[1]"));
			String DialShape_value=DialShape.getText();
			sh.getRow(i).createCell(4).setCellValue(DialShape_value);
			
			
			WebElement Water_Resistant=driver.findElement(By.xpath("//td[@class='_3-wDH3 col col-3-12'][contains (text(), 'Water Resistant')]//following::td[1]"));
			String Strap_value=Water_Resistant.getText();
			sh.getRow(i).createCell(5).setCellValue(Strap_value);

			driver.close();
			driver.switchTo().window(old_window);

		}
		
//		Logout the Flipkart
		WebElement profile = driver.findElement(By.cssSelector("#container > div > div._3ybBIU > div._1tz-RS > div._3pNZKl > div:nth-child(3) > div > div > div > div"));
		builder.moveToElement(profile).build().perform();


		WebElement logout = driver.findElement(By.linkText("Logout"));
		logout.click();

		
//		Save the Excel
		FileOutputStream fos=new FileOutputStream(file);
		wb.write(fos);
		wb.close();

      // WebElement listofProducts=driver.findElement(By.xpath("//*[text()='Apple Watch Series 4 GPS + Cellular 44 mm Space Grey Aluminium Case with Black Sport Band']"));
     //  builder.click(listofProducts).build().perform();

	}

}
