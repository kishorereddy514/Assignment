package ca.aeq.automation.sauce.pages;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import ca.aeq.automation.sauce.base.TestBase;

public class HomePage extends TestBase {

	@FindBy(xpath = "//div[@class='product_class']")
	WebElement productsLabel;
	@FindBy(xpath = "//img[@class='inventory_item_img']")
	WebElement inventoryItemImage;

	// WebElement SEL_ELEMENT = "birthday_month";
	public HomePage() {
		PageFactory.initElements(driver, this);
	}

	public String verifyHomePageTitle() {
		return driver.getTitle();
	}

	public String getProductLabelText() {
		return productsLabel.getText();
	}

	/* checks all inventory image links */
	public boolean isAllInventoryImagesClickable () throws IOException {
		boolean isDisplayed = true;
		List<WebElement> list = driver.findElements(By.xpath("//img[@class='inventory_item_img']"));
		for (int i = 0; i < list.size(); i++) {
			System.out.println(list.get(i).getAttribute("src"));
			HttpURLConnection connection = (HttpURLConnection) new URL(list.get(i).getAttribute("src"))
					.openConnection();
			connection.connect();
			String message = connection.getResponseMessage();
			connection.disconnect();
			//ToDO use log4j 
			System.out.println(list.get(i).getAttribute("src") + "----->>" + message + "response code is "
					+ connection.getResponseCode());
			if (!message.equalsIgnoreCase("Not Found")) {
				continue;
			} else {
				isDisplayed = false;
				break;
			}
		}
		return isDisplayed;

	}

	// find all broken links in home page
	public void findAllBrokenLinks() throws MalformedURLException, IOException {

		List<WebElement> list = driver.findElements(By.tagName("a")); //find by tagName 'a'
		list.addAll(driver.findElements(By.tagName("img"))); //find by tagName 'img'

		System.out.println("list of all a and img " + list.size());
		List<WebElement> activeList = new ArrayList<>();
		for (int i = 0; i < list.size(); i++) {
			System.out.println(list.get(i).getAttribute("href"));
			if (list.get(i).getAttribute("href") != null
					&& !(list.get(i).getAttribute("href").contains("javascript"))) {
				activeList.add(list.get(i)); // get all active list
			}
		}

		System.out.println("list of active href's" + activeList.size());

		for (int j = 0; j < activeList.size(); j++) {
			HttpURLConnection connection = (HttpURLConnection) new URL(activeList.get(j).getAttribute("href"))
					.openConnection();
			connection.connect();
			String message = connection.getResponseMessage(); // or check connection.getResponseCode() 
			connection.disconnect();
			System.out.println(activeList.get(j).getAttribute("href") + "----->>" + message); 
		}

	}

}
