package com.orangehrm.library;

import org.openqa.selenium.By;
import org.openqa.selenium.interactions.Actions;

import com.orangehrm.constants.Orangehrmconstants;

public class AdminHomePage extends Orangehrmconstants
{
	public boolean adminLogout()
	
	{
		String pageurl;
		driver.findElement(By.partialLinkText("Welcome")).click();
		driver.findElement(By.linkText("Logout")).click();
		pageurl=driver.getCurrentUrl();
		if (pageurl.toLowerCase().contains("login")) 
		{
			return true;
		} else 
		{
			return false;
		}
	}
	
	public void gotoEmpRegPage()
	{
		driver.findElement(By.linkText("PIM")).click();
		driver.findElement(By.linkText("Add Employee")).click();
	}
	
	public void gotoUserRegPage()
	{
	driver.findElement(By.linkText("Admin")).click();
	Actions act=new Actions(driver);
	act.moveToElement(driver.findElement(By.linkText("User Management")));
	act.build().perform();
	driver.findElement(By.linkText("Users")).click();		
	}
}
