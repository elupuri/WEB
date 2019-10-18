package com.orangehrm.library;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.orangehrm.constants.Orangehrmconstants;

public class User extends Orangehrmconstants
{

	
	public boolean addUser(String empname,String uname,String pword)
	{
		String expval,actval;
		expval=uname;
		driver.findElement(By.id("btnAdd")).click();
		driver.findElement(By.id("systemUser_employeeName_empName"))
												.sendKeys(empname);
		driver.findElement(By.id("systemUser_userName"))
												.sendKeys(uname);		
		driver.findElement(By.id("systemUser_password"))
												.sendKeys(pword);		
		driver.findElement(By.id("systemUser_confirmPassword"))
												 .sendKeys(pword);
		driver.findElement(By.id("btnSave")).click();		
		WebElement usertable;
		usertable=driver.findElement(By.id("resultTable"));
		List<WebElement> rows,cols;
		rows=usertable.findElements(By.tagName("tr"));		
		boolean res=false;		
		for (int i = 1; i < rows.size(); i++) 
		{
			cols=rows.get(i).findElements(By.tagName("td"));
			actval=cols.get(1).getText();
			if (expval.equalsIgnoreCase(actval))
			{
				res=true;
				break;
			}
		}
		return res;		
	}
	
	
	
	
	
}
