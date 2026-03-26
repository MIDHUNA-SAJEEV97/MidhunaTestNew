package testPackSample;

import com.aventstack.extentreports.*;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;

public class ReportManager {

    private static ExtentReports extent;

    public static ExtentReports getInstance() {
        if (extent == null) {

            ExtentSparkReporter spark = new ExtentSparkReporter("ExtentReport.html");
            spark.config().setDocumentTitle("Automation Test Report");
            spark.config().setReportName("Execution Summary");

            extent = new ExtentReports();
            extent.attachReporter(spark);

            extent.setSystemInfo("Executed By", System.getProperty("user.name"));
            extent.setSystemInfo("Environment", "Automation QA");
        }
        return extent;
    }
}
