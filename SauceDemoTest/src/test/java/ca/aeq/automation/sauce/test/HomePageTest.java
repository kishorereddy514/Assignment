package ca.aeq.automation.sauce.test;

import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import ca.aeq.automation.sauce.base.TestBase;
import ca.aeq.automation.sauce.pages.HomePage;
import ca.aeq.automation.sauce.pages.LoginPage;
import ca.aeq.automation.sauce.util.TestUtil;

public class HomePageTest extends TestBase {

	LoginPage loginPage;
	HomePage homePage;
	TestUtil testUtil;

	public HomePageTest() {
		super();
	}

	@BeforeMethod
	@Parameters("browser")
	public void setUp(String browser) {
		initialization(browser);
		loginPage = new LoginPage();
		testUtil = new TestUtil();
		// homePage= loginPage.loginAndContinue(prop.getProperty("userName"), prop.getProperty("password"));

		homePage = loginPage.loginAndContinue(prop.getProperty("problemUserName"), prop.getProperty("password"));
	}

	@Test
	public void verifyHomePageTitleTest() {
		String title = homePage.verifyHomePageTitle();
		Assert.assertEquals(title, prop.getProperty("homePageTitle"));
	}

	@Test
	public void verifyInventoryItemImages() {

		try {
			Assert.assertTrue(homePage.isAllInventoryImagesClickable(), "Inventory Item Iamges are not displayed");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@AfterMethod
	public void tearDown() {
		driver.quit();
	}
}
