package ca.aeq.automation.sauce.base;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.events.EventFiringWebDriver;

import ca.aeq.automation.sauce.util.ScreenShotHelper;
import ca.aeq.automation.sauce.util.TestUtil;
import ca.aeq.automation.sauce.util.WebEventListner;



public class TestBase {

	public static WebDriver driver;
	public static Properties prop;
	public static EventFiringWebDriver e_driver;
	public static WebEventListner eventListener;

	public TestBase() {
		prop = new Properties();
		try {
			FileInputStream inStream = new FileInputStream(
					System.getProperty("user.dir") + "/src/main/java/ca/aeq/automation/sauce/config/config.properties");
			prop.load(inStream);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException io) {
			io.printStackTrace();
		}
	}

	
	public static void initialization(String browser) {
		//String browser = prop.getProperty("browser");
		if (browser.equals("chrome")) {
			System.setProperty("webdriver.chrome.driver",
					System.getProperty("user.dir") + "/drivers/chromedriver/chromedriver.exe");
			driver = new ChromeDriver();
		} else if (browser.equals("FF")) {
			System.setProperty("webdriver.gecko.driver",
					System.getProperty("user.dir") + "/drivers/geckodriver/geckodriver.exe");
			driver = new FirefoxDriver();
		}
		e_driver = new EventFiringWebDriver(driver);
		eventListener = new WebEventListner();
		e_driver.register(eventListener);
		driver = e_driver;
		
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().pageLoadTimeout(TestUtil.PAGE_LOAD_TIMEOUT, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(TestUtil.IMPLICIT_WAIT, TimeUnit.SECONDS);
		
		driver.get(prop.getProperty("url"));
	}
	
	/** COMMON PAGE METHODS **/
	
	public boolean isFieldDisplayedAndEditable(WebElement element) {
		return element.isDisplayed() && element.isEnabled();
	}
	public boolean doesElementExistsById(String elementId) {
		return doesElementExists(By.id(elementId));
	}
	
	public boolean doesElementExists(By selector) {
		try {
			driver.findElement(selector);
			return true;
		} catch (NoSuchElementException e) {
			return false;
		}
	}
	
	public WebElement getElement(By selector) {
		WebElement element = null;

		try {
			element = driver.findElement(selector);
		} catch (NoSuchElementException e) {
			ScreenShotHelper.takeScreenShot(driver, System.getProperty("user.dir"),this.getClass().toString());
		}

		return element;
	}
	public boolean doesTextExist(String searchString) {
		String searchXpath = "//div[contains(text(),'" + searchString + "')]";
		return doesElementExists(By.xpath(searchXpath));
	}

	public boolean doesPageIDExist(String pageID) {
		try {
			String searchXpath = "//div[contains(@id, '" + pageID + "')]";
			return doesElementExists(By.xpath(searchXpath));
		} catch (Exception e) {
			return false;
		}
	}
	
	
	public WebElement getElementById(String elementId) {
		return getElement(By.id(elementId));
	}	

	public WebElement getElementByXPath(String elementXPath) {
		return getElement(By.xpath(elementXPath));
	}	

	public void sleepSafely(final long timeToSleep) {
		try {
			Thread.sleep(timeToSleep); 
										
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

}
