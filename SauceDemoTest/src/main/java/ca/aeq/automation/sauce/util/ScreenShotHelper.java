package ca.aeq.automation.sauce.util;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class ScreenShotHelper {
		

	public static String takeScreenShot(WebDriver webDriver, String outputDirectory, String name) {
		
		File scrFile;
		
			scrFile = ((TakesScreenshot) webDriver)
				.getScreenshotAs(OutputType.FILE);
		// Now you can do whatever you need to do with it, for example copy
		// somewhere
		try {
			String fileName = name+"_"+getDate()+"_"+getTime()+".jpg";
			FileUtils.copyFile(scrFile, new File(outputDirectory, fileName));
			return fileName;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return "Error Creating Screenshot";
		}

	}
	private  final static String getDate(  )   {  
        DateFormat df = new SimpleDateFormat( "yyyy-MM-dd" ) ;  
        df.setTimeZone( TimeZone.getTimeZone( "PST" )  ) ;  
        return ( df.format( new Date(  )  )  ) ;  
    }  
	private  final static String getTime(  )   {  
        DateFormat df = new SimpleDateFormat( "hh-mm-ss" ) ;  
        df.setTimeZone ( TimeZone.getTimeZone ( "PST" )  ) ;  
        //df.setTimeZone( TimeZone.getTimeZone( "Singapore" )  ) ;  
          
          
        return ( df.format( new Date(  )  )  ) ;  
    }  

}
