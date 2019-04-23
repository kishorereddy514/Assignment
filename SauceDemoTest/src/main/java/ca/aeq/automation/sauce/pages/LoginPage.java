package ca.aeq.automation.sauce.pages;


import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import ca.aeq.automation.sauce.base.TestBase;

public class LoginPage extends TestBase {

	// page factory or object repository

	@FindBy(id = "user-name")
	WebElement userName;

	@FindBy(id = "password")
	WebElement password;

	@FindBy(xpath = "//input[@class = 'btn_action' and @type = 'submit']")
	WebElement loginButton;

	@FindBy(xpath = "//div[contains(@class,'login_logo')]")
	WebElement pageLogo;

	@FindBy(xpath = "//div[contains(@class,'bot_column')]")
	WebElement botColumn;

	@FindBy(xpath = "//h3[contains(@data-test,'error')]")
	WebElement errorMessage;

	// Initializing page objects
	public LoginPage() {
		PageFactory.initElements(driver, this);
	}

	// Actions
	
	public String validateLoginPageTitle() {
		return driver.getTitle();
	}

	public boolean validateLoginLogo() {
		return pageLogo.isDisplayed();
	}

	public boolean validateBotColumnLogo() {
		return botColumn.isDisplayed();
	}

	public boolean isUserNameFieldDisplayedAndEditable() {
		return isFieldDisplayedAndEditable(userName);
	}

	public boolean isPasswordFieldDisplayedAndEditable() {
		return isFieldDisplayedAndEditable(password);
	}

	public void login(String un, String psw) {
		userName.sendKeys(un);
		password.sendKeys(psw);
		loginButton.click();
	}
	public HomePage loginAndContinue(String un, String psw) {
		userName.sendKeys(un);
		password.sendKeys(psw);
		loginButton.click();
		return new HomePage();
	}
	public boolean isErrorMessageDisplayed() {
		return errorMessage.isDisplayed();
	}

	public String getErrorMessage() {
		return errorMessage.getText().trim();
	}

	public boolean isCurrentPage() {
		return doesElementExistsById("user-name");
	}

}