import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class yatraAutomation {

    WebDriver driver;
    WebDriverWait wait;
    ExtentReports extent;
    ExtentTest test;
    String projectDir;

    @BeforeMethod
    @Parameters("browser")
    public void setup(String browser) {
        // Initialize ExtentSparkReporter
        ExtentSparkReporter sparkReporter = new ExtentSparkReporter("./src/test/java/reports/YatraTestReport.html");
        sparkReporter.config().setTheme(Theme.STANDARD);
        sparkReporter.config().setDocumentTitle("Yatra Test Automation Report");
        sparkReporter.config().setReportName("Yatra Test Report");

        projectDir = System.getProperty("user.dir");
        System.out.println("Project Directory: " + projectDir);

        // Initialize ExtentReports and attach SparkReporter
        extent = new ExtentReports();
        extent.attachReporter(sparkReporter);
        extent.setSystemInfo("Tester", "G Santhosh Kumar");

        // Initialize WebDriver based on the browser parameter
        if (browser.equalsIgnoreCase("chrome")) {
            driver = new ChromeDriver();
        } else if (browser.equalsIgnoreCase("ie")) {
            driver = new EdgeDriver();
        }

        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.manage().window().maximize();

        test = extent.createTest("Yatra Test - " + browser);
        test.info("Browser " + browser + " is opened.");
    }

    @Test
    public void yatraTestWithExcelData() throws IOException {
        // Read URL from Excel file
        String excelFilePath = "./src/test/resources/testdata/YatraTestData.xlsx";
        String url = getExcelData(excelFilePath, "Sheet1", 0, 0);

        // Open Yatra website
        driver.get(url);
        test.info("Website opened: " + url);

        // Perform validations and actions
        clickElement(By.linkText("Offers"), "Offers link clicked");

        validateTitle("Domestic Flights Offers | Deals on Domestic Flight Booking | Yatra.com");

        validateBannerLogo(By.xpath("//h2[contains(text(),'Great Offers & Amazing Deals')]"));

        captureScreenshot("./src/test/screenshots/s1.jpg");

        listHolidayPackages();
    }

    @AfterMethod
    public void tearDown() {
        if (driver != null) {
            driver.quit();
            test.info("Browser closed.");
        }
        extent.flush();
    }

    // ******************** Reusable Methods **************************

    // Method to click an element
    private void clickElement(By locator, String description) {
        WebElement element = wait.until(ExpectedConditions.elementToBeClickable(locator));
        element.click();
        test.info(description);
    }

    // Method to validate page title
    private void validateTitle(String expectedTitle) {
        String actualTitle = driver.getTitle();
        try {
            Assert.assertEquals(actualTitle, expectedTitle, "Title mismatch!");
            test.pass("Title validation passed: " + actualTitle);
        } catch (AssertionError e) {
            test.fail("Title validation failed! Expected: " + expectedTitle + ", Got: " + actualTitle);
        }
    }

    // Method to validate banner logo text
    private void validateBannerLogo(By locator) {
        WebElement bannerLogo = wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
        String actualLogoText = bannerLogo.getText();
        try {
            Assert.assertTrue(actualLogoText.contains("Great Offers & Amazing Deals"), "Logo mismatch!");
            test.pass("Banner logo validation passed: " + actualLogoText);
        } catch (AssertionError e) {
            test.fail("Banner logo validation failed! Got: " + actualLogoText);
        }
    }

    // Method to capture screenshot
    private void captureScreenshot(String filePath) throws IOException {
        File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        FileHandler.copy(screenshot, new File(filePath));
        test.addScreenCaptureFromPath(filePath);
        test.info("Screenshot captured: " + filePath);
    }

    // Method to list holiday packages
    private void listHolidayPackages() {
        List<WebElement> packages = driver.findElements(By.xpath("//div[@class='packageDetails']"));
        for (int i = 0; i < Math.min(5, packages.size()); i++) {
            WebElement packageDetails = packages.get(i);
            String packageName = packageDetails.findElement(By.xpath(".//h3")).getText();
            String packagePrice = packageDetails.findElement(By.xpath(".//span[contains(@class, 'price')]")).getText();
            test.info("Package " + (i + 1) + ": " + packageName + " - Price: " + packagePrice);
        }
    }

    // Method to get data from Excel file
    private String getExcelData(String filePath, String sheetName, int rowNum, int colNum) {
        String cellData = null;
        try (FileInputStream fis = new FileInputStream(filePath)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheet(sheetName);
            XSSFRow row = sheet.getRow(rowNum);
            XSSFCell cell = row.getCell(colNum);
            cellData = cell.getStringCellValue();
        } catch (IOException e) {
            test.fail("Failed to read Excel data: " + e.getMessage());
        }
        return cellData;
    }
}
