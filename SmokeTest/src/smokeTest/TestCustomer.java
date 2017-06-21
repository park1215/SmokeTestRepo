package smokeTest;
import org.testng.annotations.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.*;
import org.yaml.snakeyaml.tokens.Token.ID;
import org.testng.Assert;

public class TestCustomer {

	//public String baseUrl = "https://pgui-ooe02.test.wdc1.wildblue.net:8443/ooe/";
	
	//Support Portal
	public String baseUrl = "https://ordermgmt.test.exede.net/PublicGUI-SupportGUI/v1/login.xhtml"; 
	
	public String driverPath = "C:\\Selenium\\geckodriver.exe";
	public static WebDriver driver;
	
//	 @BeforeTest
//     public void launchBrowser() {
//         System.out.println("launching firefox browser"); 
//         System.setProperty("webdriver.firefox.marionette", driverPath);
//         driver = new FirefoxDriver();
//         driver.get(baseUrl);
//     }
	
	@Test(priority=0, enabled=true)
	public void addCustomer() throws InterruptedException 
	{
		System.out.println("launching firefox browser : SmokeTest");
		
		System.setProperty("webdriver.gecko.driver", "C:\\Selenium\\geckodriver.exe");
		
		driver = new FirefoxDriver();
		
		driver.manage().window().maximize();
		
		driver.get("http://www.viasat.com");
		
		String title = driver.getTitle();
		
		Assert.assertTrue(title.contains("ViaSat"));
//		driver.get(baseUrl);
//		
//		WebElement userNameField = driver.findElement(By.xpath("//*[@id=\"document:body\"]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td[2]/input"));
//		
//		userNameField.sendKeys("bmiller");
//		
//		WebElement passwordField = driver.findElement(By.xpath("//*[@id=\"document:body\"]/table/tbody/tr[2]/td/form/table/tbody/tr[4]/td[2]/input"));
//		
//		passwordField.sendKeys("Viasat12");
//		
//		WebElement loginButton = driver.findElement(By.xpath("//*[@id=\"document:body\"]/table/tbody/tr[2]/td/form/table/tbody/tr[5]/td[1]/input"));
//		
//		loginButton.click();
//		//WebElement submit = driver.findElement(By.xpath("//*[@id=\"tsf\"]/div[2]/div[3]/center/input[1]"));
//		
//		Thread.sleep(3000);
//		
//		System.out.println(driver.getTitle());
//		
//		WebElement addCustomerTab = driver.findElement(By.id("add"));
//		
//		addCustomerTab.click();
//		
//		Thread.sleep(2000);
//		
//		Select salesChannel = new Select(driver.findElement(By.id("addCustomerForm:salesChannelMenu")));
//		
//		salesChannel.selectByValue("WB_DIRECT");
//		
//		WebElement transactionReference = driver.findElement(By.id("addCustomerForm:transactionReference"));
//		
//		transactionReference.sendKeys("Transaction Reference");
//		
//		Thread.sleep(3000);
//		
//		Select marketingSource = new Select(driver.findElement(By.id("addCustomerForm:marketingSourceMenu")));
//		
//		marketingSource.selectByValue("QWESTREFERRAL");
//		
//		WebElement referralSource = driver.findElement(By.id("addCustomerForm:referralSource"));
//		
//		referralSource.sendKeys("Referral Source");
//		
//		Thread.sleep(2000);
//				
//		WebElement firstName = driver.findElement(By.id("addCustomerForm:namesIdName1"));
//		
//		firstName.sendKeys("Donald");
//		
//		WebElement lastName = driver.findElement(By.id("addCustomerForm:namesIdName3"));
//		
//		lastName.sendKeys("Trump");
//		
//		WebElement addressLine1 = driver.findElement(By.id("addCustomerForm:addressIdMaybeTableAddress1"));
//		
//		addressLine1.sendKeys("11828 E Maplewood Ave");
//		
//		WebElement city = driver.findElement(By.id("addCustomerForm:addressIdMaybeTableCity"));
//		
//		city.sendKeys("Englewood");
//		
//		Select state = new Select(driver.findElement(By.id("addCustomerForm:addressIdMaybeTableStateAddressState")));
//		
//		state.selectByValue("CO");
//		
//		WebElement zipCode = driver.findElement(By.id("addCustomerForm:addressIdMaybeTableZip"));
//		
//		zipCode.sendKeys("80111");
//		
//		WebElement primaryPhone = driver.findElement(By.id("addCustomerForm:primaryPhoneIdMaybeTablePhoneNumber"));
//		
//		primaryPhone.sendKeys("303 334 7453");
//		
//		WebElement noEmailAddress = driver.findElement(By.id("addCustomerForm:noEmailAddressSelectID"));
//		
//		noEmailAddress.click();
//		
//		WebElement birthday = driver.findElement(By.id("addCustomerForm:Birthdate"));
//		
//		birthday.sendKeys("01/01/1980");
//		
//		Thread.sleep(3000);
//		
//		WebElement nextButton = driver.findElement(By.id("addCustomerForm:nextButtonId"));
//		
//		Thread.sleep(3000);
//		
//		nextButton.click();
//		
//		Thread.sleep(10000);
//		
//		System.out.println(driver.getTitle());
//		
//		if(driver.getTitle() != null)
//		{
//			System.out.println("There is a page title");
//		}
//		
//		System.out.println("contacts tab displayed");
//				
//		//driver.quit();
//	
//		Thread.sleep(3000);
//		
//		nextButton = driver.findElement(By.id("addCustomerForm:nextButtonId"));
//		
//		nextButton.click();
//		
//		Thread.sleep(3000);
//
//		WebElement liberty50 = driver.findElement(By.id("addCustomerForm:topPackages:_5"));
//		
//		liberty50.click();
//		
//		Thread.sleep(10000);
//		
//		nextButton = driver.findElement(By.id("addCustomerForm:nextButtonId"));
//		
//		nextButton.click();
//			
//		Thread.sleep(3000);
//		
//		WebElement equipmentLeaseMonthly = driver.findElement(By.id("addCustomerForm:_1selectionPackages:_0"));
//		
//		equipmentLeaseMonthly.click();
//		
//		Thread.sleep(3000);
//		
//		nextButton = driver.findElement(By.id("addCustomerForm:nextButtonId"));
//		
//		nextButton.click();
//		
//		Thread.sleep(10000);
//
//		Select paymentMethod = new Select(driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableselectPaymentTypeChoiceId")));
//		
//		paymentMethod.selectByValue("CREDIT_CARD_RECURRING_PAYMENT");
//				
//		Select creditCardType = new Select(driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdcreditCardTypeId")));
//		
//		creditCardType.selectByValue("VISA");
//		
//		WebElement ccNumber = driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdNumberId"));
//		
//		ccNumber.sendKeys("4012000077777777");
//				
//		Select expirationMonth = new Select(driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdExpireMonthIdMonthId")));
//		
//		expirationMonth.selectByValue("04");
//		
//		Select expirationYear = new Select(driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdExpireYearIdYearId")));
//		
//		expirationYear.selectByValue("2019");
//	
//		WebElement firstNameOnCard = driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdFirstNameId"));
//		
//		firstNameOnCard.sendKeys("VISA");
//		
//		WebElement lastNameOnCard = driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdLastNameId"));
//		
//		lastNameOnCard.sendKeys("APPROVAL");
//		
//		WebElement ccZipCode = driver.findElement(By.id("addCustomerForm:recurringPaymentIdRecurringPaymentMethodIdTableCreditCardIdAddressZip"));
//		
//		ccZipCode.sendKeys("80111");
//		
//		Select taxJurisdiction = new Select(driver.findElement(By.id("addCustomerForm:taxJurisdictionMenu")));
//		
//		taxJurisdiction.selectByIndex(1);
//		
//		Thread.sleep(3000);
//		
//		nextButton = driver.findElement(By.id("addCustomerForm:nextButtonId"));
//		
//		nextButton.click();
//		
//		Thread.sleep(3000);
//		
//		WebElement schedule = driver.findElement(By.id("addCustomerForm:scheduleInstallationButtonId"));
//		
//		schedule.click();
//		
//		Thread.sleep(10000);
//		
//		WebElement submitOrder = driver.findElement(By.id("addCustomerForm:submitButtonId"));
//		
//		submitOrder.click();
//		
//		Thread.sleep(10000);
//		
//		WebElement accountReference = driver.findElement(By.id("addCustomerForm:accountReference"));
//		
//		String accountReferenceNumber = accountReference.getText();
//		
//		System.out.println("Customer Account Number : "+ accountReferenceNumber);
		
		driver.manage().deleteAllCookies();
		
	}
    
//	@BeforeMethod
//    public void verifyHomepageTitle() {
//        String expectedTitle = "Welcome: Mercury Tours";
//        String actualTitle = driver.getTitle();
//        Assert.assertEquals(actualTitle, expectedTitle);
//    }
//    @AfterMethod
//    public void goBackToHomepage ( ) {
//          driver.findElement(By.linkText("Home")).click() ;
//    }
	
//    @AfterTest
//    public void terminateBrowser()
//    {
//        driver.close();
//    }
}
