package ca.aeq.automation.sauce.util;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import ca.aeq.automation.sauce.base.TestBase;


public class TestUtil extends TestBase {

	public static long PAGE_LOAD_TIMEOUT = 20;
	public static long IMPLICIT_WAIT = 10;
	private static final String EMPTY_FIELD_MARKER = "none";
	static Workbook book;
	static  Sheet sheet;

	public static String TEST_SHEET_PATH = System.getProperty("user.dir") +"/src/main/java/ca/aeq/automation/sauce/testData/LoginCredentials.xlsx";
	public void swithToFrame() {
		driver.switchTo().frame("mainpanel");
	}
	
	public static void takeScreenshotAtEndOfTest() {
		try {
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		String currentDir = System.getProperty("user.dir");
		FileUtils.copyFile(scrFile, new File(currentDir + "/screenshots/" + System.currentTimeMillis() + ".png"));
		}catch (IOException e) {
			e.printStackTrace();
		}
	}

	
	public static Object[][] getTestData(String sheetName){
		FileInputStream file = null;
		try {
			file = new FileInputStream(TEST_SHEET_PATH);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			book = WorkbookFactory.create(file);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		sheet = book.getSheet(sheetName);
		Object[][] data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for(int i = 0; i<sheet.getLastRowNum();i++) {
			for(int j=0;j<sheet.getRow(0).getLastCellNum();j++) {
				data[i][j] = sheet.getRow(i+1).getCell(j).toString();
				//System.out.println(data[i][j]);
			}
		}
		return data;
		
	}
	
	
	/**
	 * Returns you the WebDriver this page is using
	 * 
	 * @return WebDriver
	 */

	public WebDriver getWebDriver() {
		return driver;
	}
	 
	public boolean doesElementExists(By selector) {
		try {
			getWebDriver().findElement(selector);
			return true;
		} catch (NoSuchElementException e) {
			return false;
		}
	}
	
	/**
	 * Waits for a certain time 
	 * @param timeToSleep the time to wait in milliseconds
	 */
	public void sleepSafely(final long timeToSleep) {
		try {
			Thread.sleep(timeToSleep); 
										
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
	
	public void scrollToElement(final WebElement element) {
		JavascriptExecutor jsEx = (JavascriptExecutor)getWebDriver();
		jsEx.executeScript("window.scrollTo(250,(arguments[0].getBoundingClientRect()).top - 100);", element);
		sleepSafely(1000);
	}
	
	/**
	 * Types the value of keys to the field specified by the By selector
	 * parameter.
	 * 
	 * @param keys
	 *            The String that you want to enter to the field
	 * 
	 * @param selector
	 *            How you will identify the field to enter the value.
	 * @throws IOException 
	 */
	public void typeField(String keys, By selector) throws IOException {

		final WebElement inputField = getElementAndClear(selector);

		scrollToElement(inputField);
		
		if (!keys.equals(EMPTY_FIELD_MARKER)) {
			inputField.sendKeys(Keys.CONTROL + "a");
			inputField.sendKeys(keys);
			sleepSafely(250);
		}
	}


	/**
	 * Gets the WebElement. Clears that element and return in.
	 * 
	 * @param selector How the field is identified
	 * @return The WebElement
	 * @throws IOException 
	 */

	public WebElement getElementAndClear(By selector) {
		WebElement element = null;

		try {
			element = getWebDriver().findElement(selector);
			element.clear();

		} catch (NoSuchElementException e) {
			takeScreenshotAtEndOfTest();
		}

		return element;
	}
	
	

	/**
	 * Gets the WebElement and return in.
	 * @param selector selector How the field is identified
	 * @return The WebElement
	 */
	public WebElement getElement(By selector) {
		WebElement element = null;

		try {
			element = getWebDriver().findElement(selector);
		} catch (NoSuchElementException e) {
			//ScreenShotHelper.takeScreenShot(webDriver, settings.getOutputDirectory(), testName, settings.isRunInGrid());
			//throw new NonTerminalTestException(e);
		}

		return element;
	}
	
}
