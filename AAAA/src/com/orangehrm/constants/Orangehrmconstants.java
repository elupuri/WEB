package com.orangehrm.constants;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

public class Orangehrmconstants 
{
	
	public	static WebDriver driver;
public	static String url="http://orangehrm.qedgetech.com";

	@BeforeSuite
	public static void launchApp()
	{
		
		System.setProperty("webdriver.gecko.driver", "C:\\Users\\welcome\\Desktop\\geckodriver.exe");
		
		driver=new FirefoxDriver();
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS);
		driver.manage().deleteAllCookies();
		driver.manage().window().maximize();
		driver.get(url); 
	}
	
	@AfterSuite
	public static void closeApp()
	{
		driver.close();
	}






}
