package com.adm.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Set;
import java.util.concurrent.TimeUnit;

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

public class UsingPoi {
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

//		Defined the sheet, name and path of the excel sheet containing the username and password 
		File file=new File("D:\\Excel\\write.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet twentyfive_products = wb.getSheetAt(1);

//		Description of Column Header  
		twentyfive_products.createRow(0).createCell(0).setCellValue("Product Description");
		twentyfive_products.getRow(0).createCell(1).setCellValue("Type");
		twentyfive_products.getRow(0).createCell(2).setCellValue("Sleeve");
		twentyfive_products.getRow(0).createCell(3).setCellValue("Fit");
		twentyfive_products.getRow(0).createCell(4).setCellValue("Fabric");
		twentyfive_products.getRow(0).createCell(5).setCellValue("Suitable For");

//		Defined the Browser and navigated to home page of Flipkart
		System.setProperty("webdriver.gecko.driver","C:\\Users\\Monica\\Selenium\\web driver\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		driver.get("https://www.flipkart.com/");
		driver.manage().window().maximize();

//		Login the Flipkart
		WebElement userid = driver.findElement(By.xpath("//input[@type='text' and @class='_2zrpKA _1dBPDZ']"));
		userid.sendKeys("monikasindhu45@gmail.com");


		WebElement password = driver.findElement(By.xpath("//input[@type='password' and @class='_2zrpKA _3v41xv _1dBPDZ']"));
		password.sendKeys("monika@06");

		WebElement login_button = driver.findElement(By.xpath("//button[@class='_2AkmmA _1LctnI _7UHT_c' and @type='submit']"));
		login_button.click();



		driver.navigate().refresh();


//		Navigate to men's T.Shirt
		WebDriverWait wait = new WebDriverWait(driver, 10);

		Actions action = new Actions(driver);
		WebElement men = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#container > div > div.zi6sUf > div > ul > li:nth-child(3) > span")));
		action.moveToElement(men).build().perform();



		WebElement men_top_wear =wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@title='T-Shirts' or contains(text(),'T-Shirts')]")));
		men_top_wear.click();


		String old_window=driver.getWindowHandle();

		String Before_Xpath = "(//div[@class='_1vC4OE'])[";
		String After_Xpath="]";

		
//		Opening 25 T.Shirt products and writing 5 product description in the Excel
		for(int i =1; i<=25; i++)
		{
			String xpath=Before_Xpath+i+After_Xpath;
			WebElement all_T_Shirts=wait.until(ExpectedConditions.elementToBeClickable(By.xpath(xpath)));
			all_T_Shirts.click();


			Set<String> allwindows=driver.getWindowHandles();
			for (String newwindow : allwindows) {
				driver.switchTo().window(newwindow);
			}

			Thread.sleep(3000);

			WebElement product_details = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='col col-11-12 ft8ug2']")));
			product_details.click();

			WebElement read_more=driver.findElement(By.xpath("//button[@class='_2AkmmA _1jdA3N']"));
			read_more.click();

			Thread.sleep(1000);

			XSSFRow row=twentyfive_products.createRow(i);

			WebElement product_info=driver.findElement(By.xpath("//span[@class='_35KyD6']"));
			String product_info_value=product_info.getText();
			twentyfive_products.getRow(i).createCell(0).setCellValue(product_info_value);


			WebElement type=driver.findElement(By.xpath("//div[contains(text(),'Type')]//following::div[1]"));
			String type_value=type.getText();
			twentyfive_products.getRow(i).createCell(1).setCellValue(type_value);


			WebElement sleeve=driver.findElement(By.xpath("//div[contains(text(),'Sleeve')]//following::div[1]"));
			String sleeve_value=sleeve.getText();
			twentyfive_products.getRow(i).createCell(2).setCellValue(sleeve_value);

			WebElement fit=driver.findElement(By.xpath("//div[contains(text(),'Fit')]//following::div[1]"));
			String fit_value=fit.getText();
			twentyfive_products.getRow(i).createCell(3).setCellValue(fit_value);

			WebElement fabric=driver.findElement(By.xpath("//div[contains(text(),'Fabric')]//following::div[1]"));
			String fabric_value=fabric.getText();
			twentyfive_products.getRow(i).createCell(4).setCellValue(fabric_value);

			WebElement suitable_For=driver.findElement(By.xpath("//div[contains(text(),'Suitable For')]//following::div[1]"));
			String suitable_For_value=suitable_For.getText();
			twentyfive_products.getRow(i).createCell(5).setCellValue(suitable_For_value);

			driver.close();
			driver.switchTo().window(old_window);

		}

		
//		Logout the Flipkart
		WebElement profile = driver.findElement(By.cssSelector("#container > div > div._3ybBIU > div._1tz-RS > div._3pNZKl > div:nth-child(3) > div > div > div > div"));
		action.moveToElement(profile).build().perform();


		WebElement logout = driver.findElement(By.linkText("Logout"));
		logout.click();

		
//		Save the Excel
		FileOutputStream fos=new FileOutputStream(file);
		wb.write(fos);
		wb.close();


	}
}
