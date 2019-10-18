package com.orangehrm.tests;

import java.io.IOException;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.orangehrm.constants.Orangehrmconstants;
import com.orangehrm.library.AdminHomePage;
import com.orangehrm.library.Employee;
import com.orangehrm.library.Login;
import com.utils.Xlutils;

public class NewEmployeeRegTest extends Orangehrmconstants
{
	String xlfile="D:\\OrangeHRMDDT\\bin\\com\\orangehrm\\testdata\\testdata.xlsx";
	String logindatasheet="AdminLoginData";
	String empdatasheet="EmpData";
	
	@Test
	public void empregTest() throws IOException
	{		
		int rc;		
		Login li=new Login();		
		li.uname=Xlutils.getCellData(xlfile, logindatasheet, 1, 0);
		li.pword=Xlutils.getCellData(xlfile, logindatasheet, 1, 1);		
		li.adminLogin(li.uname, li.pword);		
		AdminHomePage ahome=new AdminHomePage();
		rc=Xlutils.getRowCount(xlfile, empdatasheet);
		Employee emp=new Employee();		
		for (int i = 1; i <= rc; i++) 
		{
			ahome.gotoEmpRegPage();
			emp.fname=Xlutils.getCellData(xlfile, empdatasheet, i, 0);
			emp.lname=Xlutils.getCellData(xlfile, empdatasheet, i, 1);
			boolean res=emp.addEmployee(emp.fname, emp.lname);
			if (res) 
			{
				Xlutils.setCellData(xlfile, empdatasheet, i, 2, "Pass");
				Xlutils.fillGreenColor(xlfile, empdatasheet, i, 2);
			} else 
			{
				Xlutils.setCellData(xlfile, empdatasheet, i, 2, "Fail");
				Xlutils.fillRedColor(xlfile, empdatasheet, i, 2);
			}
			Assert.assertTrue(res);			
		}
		ahome.adminLogout();			
		}
}
