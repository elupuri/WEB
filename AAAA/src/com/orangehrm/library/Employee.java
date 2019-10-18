package com.orangehrm.library;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.orangehrm.constants.Orangehrmconstants;

public class Employee extends Orangehrmconstants
{

	public String fname,lname;
	
	
	public boolean addEmployee(String fname,String lname)
	{
		String expval,actval;
		expval=fname;
		
		driver.findElement(By.id("firstName")).sendKeys(fname);
		driver.findElement(By.id("lastName")).sendKeys(lname);
		driver.findElement(By.id("btnSave")).click();
		driver.findElement(By.linkText("PIM")).click();
		driver.findElement(By.linkText("Employee List")).click();
		
		WebElement emptable;		
		emptable=driver.findElement(By.id("resultTable"));
		List<WebElement> rows,cols;
		
		rows=emptable.findElements(By.tagName("tr"));
		boolean res=false;
		for (int i = 1; i <rows.size() ; i++) 
		{
			cols=rows.get(i).findElements(By.tagName("td"));
			actval=cols.get(2).getText();
			if (expval.equalsIgnoreCase(actval)) 
			{
				res=true;
				break;
			}
		}
		return res;		
	}
	
	
	
}
