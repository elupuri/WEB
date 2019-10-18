package com.orangehrm.library;

import org.openqa.selenium.By;

import com.orangehrm.constants.Orangehrmconstants;
public class Login extends Orangehrmconstants
{
	public String uname,pword;	
	
	public boolean adminLogin(String uid,String pwd)
	{
		driver.findElement(By.id("txtUsername")).sendKeys(uid);
		driver.findElement(By.id("txtPassword")).sendKeys(pwd);
		driver.findElement(By.id("btnLogin")).click();
		try 
		{
			driver.findElement(By.linkText("Welcome Admin"));
			return true;
		} catch (Exception e) 
		{
			return false;			
		}		
	}	
	
	public boolean empLogin(String fname,String uid,String pwd)
	{
		driver.findElement(By.id("txtUsername")).sendKeys(uid);
		driver.findElement(By.id("txtPassword")).sendKeys(pwd);
		driver.findElement(By.id("btnLogin")).click();
		try 
		{
			driver.findElement(By.linkText("Welcome "+fname));
			return true;
		} catch (Exception e) 
		{
			return false;			
		}		
	}	
	
	
	
	
}
