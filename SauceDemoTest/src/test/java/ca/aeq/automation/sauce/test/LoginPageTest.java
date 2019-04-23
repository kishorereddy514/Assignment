package ca.aeq.automation.sauce.test;

import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.beust.jcommander.Parameter;

import ca.aeq.automation.sauce.base.TestBase;
import ca.aeq.automation.sauce.pages.HomePage;
import ca.aeq.automation.sauce.pages.LoginPage;
import ca.aeq.automation.sauce.util.TestUtil;

public class LoginPageTest extends TestBase {
	LoginPage loginPage;
	HomePage homepage;
	TestUtil testUtil;
	String sheetName = "Sheet1";

	public LoginPageTest() {
		super();
	}

	@BeforeMethod
	@Parameters("browser")
	public void setUp(String borwser) {

		initialization(borwser);

		loginPage = new LoginPage();
	}

	@Test
	public void loginPageTest() {
		String title = loginPage.validateLoginPageTitle();
		Assert.assertEquals(title, prop.getProperty("logingPageTitle"));
	}

	@Test
	public void swagLogoImgTest() {
		boolean isDisplayed = loginPage.validateLoginLogo();
		Assert.assertTrue(isDisplayed, "Login page swag logo is not displayed");
	}

	@Test
	public void botColumnLogoImgTest() {
		boolean isDisplayed = loginPage.validateBotColumnLogo();
		Assert.assertTrue(isDisplayed, "Login page bot column logo is not displayed");
	}

	@Test
	public void isFieldsDisplayed() {
		Assert.assertTrue(
				loginPage.isUserNameFieldDisplayedAndEditable() && loginPage.isPasswordFieldDisplayedAndEditable(),
				"UserName/Password field is neither displayed or editable");
	}

	@DataProvider
	public Object[][] getLoginTestData() {
		Object[][] data = TestUtil.getTestData(sheetName);
		return data;
	}

	@Test(dataProvider = "getLoginTestData")
	public void validateTestData(String userName, String password) {
		loginPage.login(userName, password);
		if (loginPage.isCurrentPage()) {

			verifyErrorMessage(userName, password);
			Assert.assertFalse(loginPage.isCurrentPage(),
					"Error while logging with userName " + userName + " and password " + password);

		}
		
	}

	public void verifyErrorMessage(String userName, String password) {
		String errorMessage = loginPage.getErrorMessage();
		if (userName.isEmpty() && password.isEmpty()) {
			Assert.assertTrue(errorMessage.contains("Username is required"));
		} else if (userName.isEmpty()) {
			Assert.assertTrue(errorMessage.contains("Username is required"));
		} else if (password.isEmpty()) {
			Assert.assertTrue(errorMessage.contains("Password is required"));
		}//ToDo check other error cases
	}
	

	@AfterMethod
	public void tearDown() {
		driver.quit();
	}
}
