package TestNGproject;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class sample {
	public void setup() {
		System.setProperty("webdriver.gecko.driver","C:\\Users\\Monica\\Selenium\\web driver\\geckodriver.exe");
		 WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
		String url = "https://www.sathya.in/";
		driver.get(url);
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		
	}
public static void main(String[] args) {
	
}
}
