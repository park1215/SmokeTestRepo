
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

/**
 * @author Sean Park
 * This factory class processes login and returns a browser driver object depending on the input parameters from properties file.
 * It is preferred to have parameters in the properties file, to avoid hard coded values in source code.
 */

public class BrowserFactory 
{
	static WebDriver driver;
	static String driverPath = ".\\resources\\drivers\\";
	static String propertiesPath = ".\\resources\\centcom\\centcom.properties\\";
	
	public static WebDriver startBrowser(String browser, String url) throws IOException
	{
		if (browser.equalsIgnoreCase("firefox"))
		{
			System.out.println("*******************");
			System.out.println("Launching Firefox browser");
			//System.setProperty("webdriver.firefox.marionette", "C:\\Selenium\\geckodriver.exe");
			System.setProperty("webdriver.gecko.driver", driverPath+"geckodriver.exe");
			driver = new FirefoxDriver();
		}

//		else if (prop.getProperty("browser").equalsIgnoreCase("chrome"))
		else if (browser.equalsIgnoreCase("chrome"))
		{
			System.out.println("*******************");
			System.out.println("Launching Chrome browser");
			System.setProperty("webdriver.chrome.driver", driverPath+"chromedriver.exe");
			driver = new ChromeDriver();
		}
		else 
		{
			System.out.println("*******************");
			System.out.println("Launching IE browser");
			System.setProperty("webdriver.ie.driver", driverPath+"IEDriverServer.exe");
			//On IE 7 or higher on Windows Vista or Windows 7, you must set the Protected Mode settings for each zone to be the same value.
			//This is a possible workaround for protected Mode
			//*********************************************************************************************************
			DesiredCapabilities ieCapabilities = DesiredCapabilities.internetExplorer();
			ieCapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
			ieCapabilities.setCapability("ensureCleanSession", true);
			//**********************************************************************************************************
			driver = new InternetExplorerDriver();
		}

		driver.manage().window().maximize();

		driver.get(url);
		
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		return driver;
	}
}