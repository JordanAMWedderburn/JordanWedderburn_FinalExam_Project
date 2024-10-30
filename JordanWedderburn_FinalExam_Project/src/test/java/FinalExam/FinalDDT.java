package FinalExam;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;


public class FinalDDT  {
	
	public WebDriver driver;
	
	File file = new File("C:\\Users\\Wedde\\eclipse-workspace\\JordanWedderburn_FinalExam_Project\\Excel File\\Jordan Wedderburn_FinalExam_ProjectExcel.xlsx");
	XSSFWorkbook wb = new XSSFWorkbook();
	XSSFSheet sh1 = wb.createSheet("Saint-Denis-de-Brompton");
	XSSFSheet sh2 = wb.createSheet("Saint-Justin");
	XSSFSheet sh3 = wb.createSheet("Saint-Luc-de-Bellechasse");
	XSSFSheet sh4 = wb.createSheet("La Pocatière");
	XSSFSheet sh5 = wb.createSheet("La Vallée-de-l'Or");
	
	public static ExtentSparkReporter sparkReporter;
	public static ExtentReports extent;
	public static ExtentTest test;

	public void initializer() {
		sparkReporter = new ExtentSparkReporter(System.getProperty("user.dir")+"/Reports/extentSparkReport.html");
		sparkReporter.config().setDocumentTitle("Automation Report");
		sparkReporter.config().setReportName("Test Execution Report");
		sparkReporter.config().setTheme(Theme.DARK);
		sparkReporter.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
		extent = new ExtentReports();
		extent.attachReporter(sparkReporter);
	}
	
	public static String CaptureScreenshot(WebDriver driver) throws IOException{
		String FileSeperator = System.getProperty("file.seperator");
		String Extent_report_path = "."+FileSeperator+"Reports";
		File Src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		String Screenshotname = "screenshot"+Math.random()+".png";
		File Dst = new File(Extent_report_path+FileSeperator+"Screenshots"+FileSeperator+Screenshotname);
		FileUtils.copyFile(Src,Dst);
		String absPath = Dst.getAbsolutePath();
		System.out.println("Absolute path is: "+absPath);
		return absPath;
	}
  @Test (priority = 1,enabled = true)
  public void Table1() throws IOException,InterruptedException { 
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	  driver.manage().window().maximize();
	  
	  String methodName = new Exception().getStackTrace()[0].getMethodName();
	  test = extent.createTest(methodName,"Extract Table Date");
	  test.log(Status.INFO,"Started Opening Webpage");
	  test.assignCategory("Regression Testing");
	  	Thread.sleep(3000);
	  	test.log(Status.INFO,"Opening Table");
	  	test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			driver.findElement(By.xpath("//a[normalize-space()='Saint-Denis-de-Brompton']")).click();
			Thread.sleep(3000);
			test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			WebElement iframeElement = driver.findElement(By.tagName("iframe"));
			driver.switchTo().frame(iframeElement);
			
			WebElement table = driver.findElement(By.xpath("//table[2]"));
			List<WebElement> totalRows = table.findElements(By.tagName("tr"));
			for(int row=0; row<totalRows.size(); row++)
			{
				XSSFRow rowValue = sh1.createRow(row);
				List<WebElement> totalColumns = totalRows.get(row).findElements(By.tagName("td"));
				for(int col=0; col<totalColumns.size(); col++)
				{
					String cellValue = totalColumns.get(col).getText();
					System.out.print(cellValue + "\t");
					rowValue.createCell(col).setCellValue(cellValue);
				}
				System.out.println();
			}
			test.log(Status.PASS,"Successfully Extracted Data from Table");
			driver.close();
			
  } 
  @Test (priority = 2,enabled = true)
  public void Table2() throws IOException,InterruptedException {
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	  driver.manage().window().maximize();
	  
	  String methodName = new Exception().getStackTrace()[0].getMethodName();
	  test = extent.createTest(methodName,"Extract Table Date");
	  test.log(Status.INFO,"Started Opening Webpage");
	  test.assignCategory("Regression Testing");
	  	Thread.sleep(3000);
	  	test.log(Status.INFO,"Opening Table");
	  	test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			driver.findElement(By.xpath("//a[normalize-space()='Saint-Justin']")).click();
			Thread.sleep(3000);
			test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			WebElement iframeElement = driver.findElement(By.tagName("iframe"));
			driver.switchTo().frame(iframeElement);
			
			WebElement table = driver.findElement(By.xpath("//table[2]"));
			List<WebElement> totalRows = table.findElements(By.tagName("tr"));
			for(int row=0; row<totalRows.size(); row++)
			{
				XSSFRow rowValue = sh2.createRow(row);
				List<WebElement> totalColumns = totalRows.get(row).findElements(By.tagName("td"));
				for(int col=0; col<totalColumns.size(); col++)
				{
					String cellValue = totalColumns.get(col).getText();
					System.out.print(cellValue + "\t");
					rowValue.createCell(col).setCellValue(cellValue);
				}
				System.out.println();
			}
			test.log(Status.PASS,"Successfully Extracted Data from Table");
			driver.close();
  }
  @Test (priority = 3,enabled = true)
  public void Table3() throws IOException,InterruptedException {
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	  driver.manage().window().maximize();
	  
	  String methodName = new Exception().getStackTrace()[0].getMethodName();
	  test = extent.createTest(methodName,"Extract Table Date");
	  test.log(Status.INFO,"Started Opening Webpage");
	  test.assignCategory("Regression Testing");
	  	Thread.sleep(3000);
	  	test.log(Status.INFO,"Opening Table");
	  	test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			driver.findElement(By.xpath("//a[normalize-space()='Saint-Luc-de-Bellechasse']")).click();
			Thread.sleep(3000);
			test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			WebElement iframeElement = driver.findElement(By.tagName("iframe"));
			driver.switchTo().frame(iframeElement);
			
			WebElement table = driver.findElement(By.xpath("//table[2]"));
			List<WebElement> totalRows = table.findElements(By.tagName("tr"));
			for(int row=0; row<totalRows.size(); row++)
			{
				XSSFRow rowValue = sh3.createRow(row);
				List<WebElement> totalColumns = totalRows.get(row).findElements(By.tagName("td"));
				for(int col=0; col<totalColumns.size(); col++)
				{
					String cellValue = totalColumns.get(col).getText();
					System.out.print(cellValue + "\t");
					rowValue.createCell(col).setCellValue(cellValue);
				}
				System.out.println();
			}
			test.log(Status.PASS,"Successfully Extracted Data from Table");
			driver.close();
 }
  @Test (priority = 4,enabled = true)
  public void Table4() throws IOException,InterruptedException {
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	  driver.manage().window().maximize();
	  
	  String methodName = new Exception().getStackTrace()[0].getMethodName();
	  test = extent.createTest(methodName,"Extract Table Date");
	  test.log(Status.INFO,"Started Opening Webpage");
	  test.assignCategory("Regression Testing");
	  	Thread.sleep(3000);
	  	test.log(Status.INFO,"Opening Table");
	  	test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			driver.findElement(By.xpath("//a[normalize-space()='La Pocatière']")).click();
			Thread.sleep(3000);
			test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			WebElement iframeElement = driver.findElement(By.tagName("iframe"));
			driver.switchTo().frame(iframeElement);
			
			WebElement table = driver.findElement(By.xpath("//table[2]"));
			List<WebElement> totalRows = table.findElements(By.tagName("tr"));
			for(int row=0; row<totalRows.size(); row++)
			{
				XSSFRow rowValue = sh4.createRow(row);
				List<WebElement> totalColumns = totalRows.get(row).findElements(By.tagName("td"));
				for(int col=0; col<totalColumns.size(); col++)
				{
					String cellValue = totalColumns.get(col).getText();
					System.out.print(cellValue + "\t");
					rowValue.createCell(col).setCellValue(cellValue);
				}
				System.out.println();
			}try {
				FileOutputStream outPutStream = new FileOutputStream (file);
				 wb.write(outPutStream);
			  }
			  catch (IOException e) {
			  e.printStackTrace();
		  }
			test.log(Status.PASS,"Successfully Extracted Data from Table");
			driver.close();
 } 
  @Test (priority = 5,enabled = true)
  public void Table5() throws IOException,InterruptedException {
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	  driver.manage().window().maximize();
	  
	  String methodName = new Exception().getStackTrace()[0].getMethodName();
	  test = extent.createTest(methodName,"Extract Table Date");
	  test.log(Status.INFO,"Started Opening Webpage");
	  test.assignCategory("Regression Testing");
	  	Thread.sleep(3000);
	  	test.log(Status.INFO,"Opening Table");
	  	test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			driver.findElement(By.xpath("/html[1]/body[1]/form[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[4]/tr[2]/td[1]/a[1]")).click();
			Thread.sleep(3000);
			test.addScreenCaptureFromPath(CaptureScreenshot(driver));
			WebElement iframeElement = driver.findElement(By.tagName("iframe"));
			driver.switchTo().frame(iframeElement);
			
			WebElement table = driver.findElement(By.xpath("//table[2]"));
			List<WebElement> totalRows = table.findElements(By.tagName("tr"));
			for(int row=0; row<totalRows.size(); row++)
			{
				XSSFRow rowValue = sh5.createRow(row);
				List<WebElement> totalColumns = totalRows.get(row).findElements(By.tagName("td"));
				for(int col=0; col<totalColumns.size(); col++)
				{
					String cellValue = totalColumns.get(col).getText();
					System.out.print(cellValue + "\t");
					rowValue.createCell(col).setCellValue(cellValue);
				}
				System.out.println();
			}try {
				FileOutputStream outPutStream = new FileOutputStream (file);
				 wb.write(outPutStream);
			  }
			  catch (IOException e) {
			  e.printStackTrace();
		  }
			test.log(Status.PASS,"Successfully Extracted Data from Table");
			driver.close();
}
  
  @BeforeTest
  public void beforeTest()  {
	  initializer();
  }

  @AfterTest
  public void afterTest() {
	  extent.flush();
  }
}
