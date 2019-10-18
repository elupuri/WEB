package com.orangehrm.tests;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.orangehrm.constants.Orangehrmconstants;
import com.orangehrm.library.AdminHomePage;
import com.orangehrm.library.Login;

public class AdminLoginTest extends Orangehrmconstants
{
	@Test
	public void adminLoginTest()
	{
		boolean res;
		Login li=new Login();
		li.uname="Admin";
		li.pword="Admin";
		res=li.adminLogin(li.uname, li.pword);
		Assert.assertTrue(res);
		
		AdminHomePage ahome=new AdminHomePage();
		//Sleeper.sleepTight(2);
		ahome.adminLogout();		
	}
	
	
	
	
	
}
