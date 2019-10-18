package com.orangehrm.tests;

import org.testng.annotations.Test;

import com.orangehrm.constants.Orangehrmconstants;
import com.orangehrm.library.AdminHomePage;
import com.orangehrm.library.Login;

public class EmployeeLoginTest extends Orangehrmconstants 
{

	@Test
	public void empLoginTest()
	{
		Login li=new Login();
		li.empLogin("sekar","sekar","demo123");
		AdminHomePage ahome=new AdminHomePage();
		ahome.adminLogout();
	}
		
	
}
