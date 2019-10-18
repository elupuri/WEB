package com.orangehrm.tests;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.orangehrm.constants.Orangehrmconstants;
import com.orangehrm.library.AdminHomePage;
import com.orangehrm.library.Login;
import com.orangehrm.library.User;

public class NewUserRegTest extends Orangehrmconstants 
{
	@Test
	public void userRegTest()
	{
		Login li=new Login();
		li.adminLogin("Admin", "Admin");
		AdminHomePage ahome=new AdminHomePage();
		ahome.gotoUserRegPage();
		User us=new User();
  boolean res=us.addUser("sekar demo", "sekar", "demo123");
		Assert.assertTrue(res);
		ahome.adminLogout();		
	}
	
	
	
	
}
