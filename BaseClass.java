package testPackSample;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;

import com.aventstack.extentreports.*;
import com.aventstack.extentreports.markuputils.*;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.lang.reflect.Method;


    public class BaseClass {

        protected WebDriver driver;   // <-- This is the SAME driver used in test class
        protected WebDriverWait wait;
        protected ExtentReports extent = ReportManager.getInstance();
        protected ExtentTest test;

        @BeforeMethod(alwaysRun = true)
        public void setUp(Method method) {

//            System.setProperty("webdriver.edge.driver",
//                    "C:\\Users\\N53815\\Downloads\\edgedriver_win64_144.0.3719.104\\msedgedriver.exe");
//// ------To run in headless mode-----------------------
//            EdgeOptions options = new EdgeOptions();
//            options.addArguments("--headless=new");
//            options.addArguments("--window-size=1920,1080");
//            options.addArguments("--disable-gpu");
//            options.addArguments("--no-sandbox");
//            driver = new EdgeDriver(options);
//--------------------------------------------------------
            driver = new EdgeDriver();
            driver.manage().window().maximize();
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

            test = extent.createTest(method.getName());
        }
        private void jsClick(WebElement ele) {
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", ele);
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", ele);
        }

        // Utility scroll method
        public WebElement scrollUntilVisible(By locator) {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
            WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
            ((JavascriptExecutor) driver)
                    .executeScript("arguments[0].scrollIntoView({block: 'center'});", element);
            return element;
        }

        @AfterMethod(alwaysRun = true)
        public void captureResult(ITestResult result) {
            if (result.getStatus() == ITestResult.FAILURE) {

                // Failure message
                test.fail(" Test Failed: " + result.getThrowable().getMessage());
                // Stacktrace block
                String stackTrace = org.apache.commons.lang3.exception.ExceptionUtils.getStackTrace(result.getThrowable());
                test.fail(MarkupHelper.createCodeBlock(stackTrace));

                // Screenshot
                String screenshotPath = takeScreenshot(result.getName());
                test.addScreenCaptureFromPath(screenshotPath);

            }
            else if (result.getStatus() == ITestResult.SUCCESS) {
                test.pass(" Test Passed");
            }
            else {
                test.skip(" Test Skipped");
            }
        }

        private String takeScreenshot(String testName) {
            try {
                String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
                File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

                File dest = new File("screenshots/" + testName + "_" + timestamp + ".png");
                dest.getParentFile().mkdirs();

                Files.copy(src.toPath(), dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
                return dest.getAbsolutePath();

            } catch (Exception e) {
                System.out.println("Screenshot failed: " + e.getMessage());
                return null;
            }
        }
        @AfterSuite(alwaysRun = true)
        public void flushReport() {
            extent.flush();
            if (driver != null) driver.quit();
        }



    }

