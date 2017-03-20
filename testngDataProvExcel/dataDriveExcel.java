package testngDataProvExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;



public class dataDriveExcel {

	public static WebDriver driver;
	public String filePath = "C:\\Selenium\\testngDataProviExcel\\src\\testngDataProvExcel\\excelData.xlsx";
	XSSFWorkbook wb;

	@BeforeTest
	public void beforetest() throws IOException {
		// ProfilesIni profile = new ProfilesIni();
		// FirefoxProfile myprof =profile.getProfile("myProfile");

		// Code for dataProv.properties mentioning browser
		Properties prop = new Properties();
		FileInputStream fis = new FileInputStream("C:\\Selenium\\testngDataProviExcel\\dataProv.properties");
		prop.load(fis);

		if (prop.getProperty("browser").equals("firefox")) {
			System.setProperty("webdriver.gecko.driver", "C:/Selenium/geckodriver-v0.13.0-win64/geckodriver.exe");
			driver = new FirefoxDriver();
			driver.manage().window().maximize();
		} else if (prop.getProperty("browser").equals("chrome")) {
			System.setProperty("webdriver.chrome.driver", "C:/Selenium/chromedriver_win32/chromedriver.exe");
			driver = new ChromeDriver();

		} else {
			System.setProperty("webdriver.ie.driver", "C:\\Selenium\\IEDriverServer_x64_3.0.0\\iedriverserver.exe");
			driver = new InternetExplorerDriver();
		}

	}

	@DataProvider(name = "loginCredentials")
	public Object[][] passData() throws Exception {
		excelDataConfig config = new excelDataConfig(filePath);
		int rows = config.getRowCount(0);
		Object[][] data = new Object[rows][2];
		for (int i = 0; i < rows; i++) {

			data[i][0] = config.getData(0, i, 0);
			data[i][1] = config.getData(0, i, 1);
		}
		return data;
	}

	@Test(dataProvider = "loginCredentials")
	public void atTest(String uname, String pswd) throws InterruptedException {
		/*
		 * driver.get("http://www.msn.com/en-au/");
		 * driver.findElement(By.xpath(".//*[@id='meCtrl']/a/span")).click();
		 * Thread.sleep(5000); driver.findElement(By.xpath(
		 * ".//div[text()='Email, phone, or Skype name']")).sendKeys(uname);
		 * driver.findElement(By.id("idSIButton9")).click(); Thread.sleep(5000);
		 * driver.findElement(By.xpath(".//div[text()='Password']")).sendKeys(
		 * pswd); driver.findElement(By.id("idChkBx_PWD_KMSI0Pwd")).click(); //
		 * click to // keep // signed in driver.findElement(By.xpath(
		 * ".//input[@value='Sign in']")).click(); }
		 */

		driver.get("https://facebook.com");
		// driver.findElement(By.xpath(".//input[@value='Log In to Existing
		// Account']")).click();
		driver.findElement(By.xpath(".//*[@id='email']")).sendKeys(uname);
		driver.findElement(By.xpath(".//*[@id='pass']")).sendKeys(pswd);
		driver.findElement(By.xpath(".//input[@value='Log In']")).click();
	}

	@AfterTest
	public void afterTest() {
		driver.quit();
	}
}