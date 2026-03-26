package testPackSample;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import utils.ExcelUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class PolicyCreatorTestClass extends BaseClass {

    WebDriverWait wait;

    @Test
    public void CreateMultipleQuote() throws IOException, InterruptedException {

        String filepath = System.getProperty("user.dir") + "\\TestResources\\UWTestDataSampleAutomation.xlsx";
        //Getting Total Row Count in the Excel sheet
        int NofRows = ExcelUtils.getRowCount(filepath, "UW Details");

        //Getting Total Col Counts in the Excel sheet
        int NofCols = ExcelUtils.getCellCount(filepath, "UW Details", NofRows);

        //Creating new word doc for Capturing evidence
        XWPFDocument doc = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(System.getProperty("user.dir") + "\\TestResources\\TestEvidence.docx");


        wait = new WebDriverWait(driver, Duration.ofSeconds(30));
       // Loop through Excel rows
        for (int r = 8; r <= 9; r++) {

            UWData data = readExcelRow(filepath, r);
            driver.get("https://ops.digital-trading.int.hub.allianz.co.uk/Public/PreQuoteQuestions?PackageId=7&pageNumber=1");

            selectBrokerAccount(data);
            fillInsuredDetails(data);
            fillCoverDates(data);
            fillDeclarations();
            fillClaimsHistory();
            fillMaterialDamage(data);
            fillAdditionalCovers();
            fillUnderwriterEditorsAndReferral();
            fillEndorseClause();
            fillQuoteSummary(data, filepath, r);
            fillPaymentDetails();
            //extractPolicyNumber();
            extractPolicyNumber(data, doc, filepath, r);

            System.out.println("Quote created using row: " + r);
        }
            // SAVE & CLOSE WORD DOCUMENT
            doc.write(out);
            out.close();
            doc.close();
            System.out.println("Evidence document saved successfully.");


    }

    // =============================================================
    // MODEL CLASS TO HOLD ONE ROW OF EXCEL DATA
    //This acts as a Data Transfer Object  to store Excel row values
    // =============================================================
    class UWData {
        String accountType, brokerName, contactName, contactPhone, contactEmail;
        String tradingStatus, companyReg, insuredName, postcode;
        String establishedDate, turnover, businessActivity, coverStartDate,totalDeclaredValueOfInstalledComputer,singleItemLossLimitForInstalledComputer;
    }

    // =============================================================
    // Read an Excel row (This uses  model class UWData)
    // =============================================================
    private UWData readExcelRow(String filepath, int r) throws IOException {
        UWData d = new UWData();
       // readExcelRow() — Read one full row from Excel

        d.accountType = ExcelUtils.getCellData(filepath, "UW Details", r, 1);
        d.brokerName = ExcelUtils.getCellData(filepath, "UW Details", r, 2);
        d.contactName = ExcelUtils.getCellData(filepath, "UW Details", r, 3);
        d.contactPhone = ExcelUtils.getCellData(filepath, "UW Details", r, 4);
        d.contactEmail = ExcelUtils.getCellData(filepath, "UW Details", r, 5);
        d.tradingStatus = ExcelUtils.getCellData(filepath, "UW Details", r, 6);
        d.companyReg = ExcelUtils.getCellData(filepath, "UW Details", r, 7);
        d.insuredName = ExcelUtils.getCellData(filepath, "UW Details", r, 8);
        d.postcode = ExcelUtils.getCellData(filepath, "UW Details", r, 9);
        d.establishedDate = ExcelUtils.getCellData(filepath, "UW Details", r, 10);
        d.turnover = ExcelUtils.getCellData(filepath, "UW Details", r, 11);
        d.businessActivity = ExcelUtils.getCellData(filepath, "UW Details", r, 12);
        d.coverStartDate = ExcelUtils.getCellData(filepath, "UW Details", r, 13);
        d.totalDeclaredValueOfInstalledComputer = ExcelUtils.getCellData(filepath, "UW Details", r, 14);
        d.singleItemLossLimitForInstalledComputer = ExcelUtils.getCellData(filepath, "UW Details", r, 15);
        return d;
    }


    private void selectBrokerAccount(UWData d) throws InterruptedException {
        WebElement OrganisationAccountName = driver.findElement(By.id("SelectedAgentGroupName"));
        OrganisationAccountName.sendKeys(d.accountType);
        WebElement OrganisationDesiredOption = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[contains(text(), '" + d.accountType + "')]")));
        OrganisationDesiredOption.click();
        Thread.sleep(500);

        WebElement BrokerAccountName = scrollUntilVisible(By.id("AssignedSalesPersonId"));
        BrokerAccountName.click();
        Thread.sleep(500);
// Create Select object for dropdown
        Select select = new Select(BrokerAccountName);
        String cleanName = d.brokerName                   // Clean Excel string (critical fix)
                .trim()
                .replace("\u00A0", "")      // remove non-breaking spaces
                .replaceAll("\\s+", " ");  // normalize multi-spaces
        System.out.println("Broker from excel: [" + cleanName + "]");

        try {
            select.selectByVisibleText(cleanName);                // Try direct select first
        } catch (Exception e) {

            // Fallback: loop through all options and match ignoring case
            for (WebElement option : select.getOptions()) {
                if (option.getText().trim().equalsIgnoreCase(cleanName)) {
                    option.click();
                    break;
                }
            }
        }
//        //extracting the list of names from dropdown for adding in excel(onetime purpose)
//        List<String> brokerNames = select.getOptions().stream()
//                .filter(o -> {
//                    String val = o.getAttribute("value");
//                    String text = o.getText().trim();
//                    return val != null && !val.isEmpty() && !text.equalsIgnoreCase("Please select...");
//                })
//                .map(o -> o.getText().trim())
//                .toList();
//        System.out.println("The list of broker names are :"+brokerNames);
//        //-----------------------------------------------------
        BrokerAccountName.click(); // Final click to close dropdown

        WebElement ContactName = driver.findElement(By.id("InsuranceAdviserName"));
        ContactName.sendKeys(d.contactName);
        WebElement ContactPhoneNumber = driver.findElement(By.id("InsuranceAdviserPhone"));
        ContactPhoneNumber.sendKeys(d.contactPhone);
        WebElement ContactEmailAddress = driver.findElement(By.id("InsuranceAdviserEmail"));
        ContactEmailAddress.sendKeys(d.contactEmail);
        WebElement AccountSelectionContinueBtn = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[2]/button[2]"));
        AccountSelectionContinueBtn.click();
    }

    private void fillInsuredDetails(UWData d) throws InterruptedException {
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/h1[2]")));
        driver.navigate().refresh();
        WebElement LegaltradingStatus = driver.findElement(By.id("LegalTradingStatus"));
        Select tradingoption = new Select(LegaltradingStatus);

        String cleanName = d.tradingStatus                   // Clean Excel string (critical fix)
                .trim()
                .replace("\u00A0", "")      // remove non-breaking spaces
                .replaceAll("\\s+", " ");  // normalize multi-spaces
        System.out.println("Legal Trading status from excel: [" + cleanName + "]");
        tradingoption.selectByVisibleText(cleanName);

        //---extracting the list of names from dropdown for adding in excel(onetime purpose)-------
//        List<String> legaltradingStatusList = tradingoption.getOptions().stream()
//                .filter(o -> {
//                    String val = o.getAttribute("value");
//                    String text = o.getText().trim();
//                    return val != null && !val.isEmpty() && !text.equalsIgnoreCase("Please select...");
//                })
//                .map(o -> o.getText().trim())
//                .toList();
//        System.out.println("The list Legal Trading status are :"+ legaltradingStatusList);
        //-----------------------------------------------------------

// Dynamic IDs according to the option selected from excel
        String insuredNameNumId = "";
        String companyRegNumId = "";
        String tradingNameRadioPrefix = "";
        String subsidiaryRadioPrefix = "";

        if (cleanName.equalsIgnoreCase("Private Limited")) {
            companyRegNumId = "StatusPrivateLimitedCompanyRegNumber";
            insuredNameNumId = "StatusPrivateLimitedInsuredName";
            tradingNameRadioPrefix = "StatusPrivateLimitedTradingNameYNNo";
            subsidiaryRadioPrefix = "StatusPrivateLimitedSubsidiaryCoveredNo";

        } else if (cleanName.equalsIgnoreCase("Public Limited")) {
            companyRegNumId = "StatusPublicLimitedCompanyRegNumber";
            insuredNameNumId = "StatusPublicLimitedInsuredName";
            tradingNameRadioPrefix = "StatusPublicLimitedTradingNameYNNo";
            subsidiaryRadioPrefix = "StatusPublicLimitedSubsidiaryCoveredNo";
        } else if (cleanName.equalsIgnoreCase("Private Unlimited")) {
            companyRegNumId = "StatusPrivateUnlimitedCompanyRegNumber";
            insuredNameNumId = "StatusPrivateUnlimitedInsuredName";
            tradingNameRadioPrefix = "StatusPrivateUnlimitedTradingNameYNNo";
            subsidiaryRadioPrefix = "StatusPrivateUnlimitedSubsidiaryCoveredNo";
        } else if (cleanName.equalsIgnoreCase("Charity")) {
            companyRegNumId = "StatusCharityCompanyRegNumber";
            insuredNameNumId = "StatusCharityInsuredName";
            tradingNameRadioPrefix = "StatusCharityTradingNameYNNo";
            subsidiaryRadioPrefix = "StatusCharitySubsidiaryCoveredNo";
        } else if (cleanName.equalsIgnoreCase("Trust")) {
            companyRegNumId = "StatusTrustCompanyRegNumber";
            insuredNameNumId = "StatusTrustInsuredName";
            tradingNameRadioPrefix = "StatusTrustTradingNameYNNo";
            subsidiaryRadioPrefix = "StatusTrustSubsidiaryCoveredNo";
        } else if (cleanName.equalsIgnoreCase("Limited Liability Partnership")) {
            companyRegNumId = "StatusLimitedLiabilityPartnershipCompanyRegNumber";
            insuredNameNumId = "StatusLimitedLiabilityPartnershipInsuredName";
            tradingNameRadioPrefix = "StatusLimitedLiabilityPartnershipTradingNameYNNo";
            subsidiaryRadioPrefix = "StatusLimitedLiabilityPartnershipSubsidiaryCoveredNo";
        } else {
            throw new RuntimeException("Invalid Trading Status");
        }
        // add company reg no: from excel
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(companyRegNumId)))
                .sendKeys(d.companyReg);
        // add company Insure name: from excel
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(insuredNameNumId)))
                .sendKeys(d.insuredName);

//---------------------------------------------------------------------------
// Click NO for both radios
        By tradingNameNo = By.xpath("//label[.//input[starts-with(@id,'" + tradingNameRadioPrefix + "') and @value='No']]");
        WebElement tradingLabel = wait.until(ExpectedConditions.presenceOfElementLocated(tradingNameNo));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", tradingLabel);

        By subsidiaryNo = By.xpath("//label[.//input[starts-with(@id,'" + subsidiaryRadioPrefix + "') and @value='No']]");
        WebElement subsidiaryLabel = wait.until(ExpectedConditions.presenceOfElementLocated(subsidiaryNo));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", subsidiaryLabel);

        //-----------------------------------------------------------------------------------------------
        WebElement Postcode = driver.findElement(By.id("CustomerAddressPostcode1023"));
        Postcode.click();
        Postcode.sendKeys(d.postcode);
        WebElement findAddress = driver.findElement(By.xpath("//*[@id=\"Lookup1023\"]/button"));
        findAddress.click();
        WebElement AddressOption = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"question1023\"]/div[4]/button[6]")));
        AddressOption.click();
        WebElement BusinessEstablishDate = driver.findElement(By.id("BusinessEstablishedDate"));
        BusinessEstablishDate.sendKeys(d.establishedDate);
        WebElement AnnualturnOver = driver.findElement(By.id("EstimatedAnnualTurnover"));
        AnnualturnOver.sendKeys(d.turnover);
        WebElement BusinessActivity = driver.findElement(By.id("Instanda_PrimaryTrade"));
        BusinessActivity.sendKeys(d.businessActivity);

        List<WebElement> options = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("Instanda_PrimaryTrade_listbox")));
        for (WebElement option : options) {
            if (option.getText().contains(d.coverStartDate)) {
                JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;
                jsExecutor.executeScript("arguments[0].click();", option);
                System.out.println("Option 'Accountancy' selected successfully.");
                WebElement desiredoption = driver.findElement(By.cssSelector(".tt-suggestion.tt-selectable,.tt-highlight"));
                desiredoption.click();
                break;
            }
        }

        WebElement EstimatedPercentageTurnover = driver.findElement(By.id("EstimatedPercentageTurnover"));
        EstimatedPercentageTurnover.clear();
        EstimatedPercentageTurnover.sendKeys("100");
        WebElement SecondaryBusinessActivity = driver.findElement(By.xpath("//*[@id=\"question1091\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement SecBusinessActivityOption = SecondaryBusinessActivity.findElement(By.xpath("//*[@id=\"question1091\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        SecBusinessActivityOption.click();
        WebElement InsuredContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        InsuredContinueButton.click();
    }


    private void fillCoverDates(UWData d) {
        WebElement PolicyInceptionDate = driver.findElement(By.id("PolicyInceptionDate"));
        PolicyInceptionDate.clear();
        PolicyInceptionDate.sendKeys(d.coverStartDate);
        WebElement PolicyEndDate = driver.findElement(By.xpath("//*[@id=\"question1082\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement PolicyEndDateOption = PolicyEndDate.findElement(By.xpath("//*[@id=\"question1082\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        PolicyEndDateOption.click();
        WebElement CoverDatesContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        CoverDatesContinueButton.click();
    }

    private void fillDeclarations() {
        WebElement CriminalConvictions = driver.findElement(By.xpath("//*[@id=\"question1722\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement CCOption = CriminalConvictions.findElement(By.xpath("//*[@id=\"question1722\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        CCOption.click();
        WebElement BankandLiquidations = driver.findElement(By.xpath("//*[@id=\"question1727\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement BankandLiquidationsOption = BankandLiquidations.findElement(By.xpath("//*[@id=\"question1727\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        BankandLiquidationsOption.click();
        WebElement InsuranceVoid = driver.findElement(By.xpath("//*[@id=\"question1731\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement InsuranceVoidOption = InsuranceVoid.findElement(By.xpath("//*[@id=\"question1731\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        InsuranceVoidOption.click();
        WebElement CCJs = driver.findElement(By.xpath("//*[@id=\"question1733\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement CCJsOption = CCJs.findElement(By.xpath("//*[@id=\"question1733\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        CCJsOption.click();
        WebElement Disqualifications = driver.findElement(By.xpath("//*[@id=\"question1746\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement DisqualificationOption = Disqualifications.findElement(By.xpath("//*[@id=\"question1746\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        DisqualificationOption.click();
        WebElement HMRevenue = driver.findElement(By.xpath("//*[@id=\"question1756\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement HMRevenueOption = HMRevenue.findElement(By.xpath("//*[@id=\"question1756\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        HMRevenueOption.click();
        WebElement HealthyandSafety = driver.findElement(By.xpath("//*[@id=\"question1758\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement HealthyandSafetyOption = HealthyandSafety.findElement(By.xpath("//*[@id=\"question1758\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        HealthyandSafetyOption.click();
        WebElement Litigation = driver.findElement(By.xpath("//*[@id=\"question1765\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement LitigationOption = Litigation.findElement(By.xpath("//*[@id=\"question1765\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        LitigationOption.click();
        WebElement DeclarationsContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        DeclarationsContinueButton.click();
    }

    private void fillClaimsHistory() {
        WebElement ClaimsHistoryContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        ClaimsHistoryContinueButton.click();
    }

    private void fillMaterialDamage(UWData d) throws InterruptedException {
        WebElement yesRadioButton = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@name='PremisesIsSameAsClientAddress__11__1' and @value='Yes']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", yesRadioButton);
        Thread.sleep(1000);
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", yesRadioButton);
        WebElement Wallconsturction = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Premises1_ConstructionTypeValidation_1Yes")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", Wallconsturction);
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", Wallconsturction);
        Thread.sleep(500);
        WebElement GroundFloor = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Premises1_InstalledComputerEquipmentOnTheGroundFloorYes")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", GroundFloor);
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", GroundFloor);
        Thread.sleep(500);
        WebElement Flooding = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Premises1_PreviouslySufferedFloodDamageNo")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", Flooding);
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", Flooding);
        Thread.sleep(500);
        WebElement InstalledComputerPremisesValue = driver.findElement(By.id("Premises1_InstalledComputerEquipmentDeclaredValue"));
        InstalledComputerPremisesValue.sendKeys(d.totalDeclaredValueOfInstalledComputer);
        WebElement InstallComputerSingleItemLimit = driver.findElement(By.id("Premises1_InstalledComputerEquipmentLossLimit"));
        InstallComputerSingleItemLimit.sendKeys(d.singleItemLossLimitForInstalledComputer);
        WebElement PortableComputer = driver.findElement(By.xpath("//*[@id=\"question2100\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement PortableComputerOption = PortableComputer.findElement(By.xpath("//*[@id=\"question2100\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        PortableComputerOption.click();
        WebElement MaterialDamageContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        MaterialDamageContinueButton.click();
    }

    private void fillAdditionalCovers() {
        WebElement ComputerMedia = driver.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement ComputerMediarOption = ComputerMedia.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        ComputerMediarOption.click();
        WebElement AdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement AdditionalExpenditureOption = AdditionalExpenditure.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        AdditionalExpenditureOption.click();
        WebElement ERisks = driver.findElement(By.xpath("//*[@id=\"question1214\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement ERisksOption = ERisks.findElement(By.xpath("//*[@id=\"question1214\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        ERisksOption.click();
        WebElement BreakdownBusiness = driver.findElement(By.xpath("//*[@id=\"question1195\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement BreakdownBusinessOption = BreakdownBusiness.findElement(By.xpath("//*[@id=\"question1195\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        BreakdownBusinessOption.click();
        WebElement Terrorism = driver.findElement(By.xpath("//*[@id=\"question1217\"]/div[2]/div[1]/div/div[1]"));
        WebElement TerrorismOption = Terrorism.findElement(By.xpath("//*[@id=\"question1217\"]/div[2]/div[1]/div/div[1]/label[2]"));
        TerrorismOption.click();
        WebElement AdditonalPageContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        AdditonalPageContinueButton.click();
    }

    private void fillUnderwriterEditorsAndReferral() throws InterruptedException {
        System.out.println("checkbox selection.");
        WebElement checkbox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("Lines_0__IsSelected")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", checkbox);
        checkbox.click();
        WebElement clearbutton = driver.findElement(By.id("clearButton"));
        clearbutton.click();
        WebElement reasontextbox = driver.findElement(By.id("Reason"));
        reasontextbox.sendKeys("cleared the referral");
        WebElement clearviewendorsementbutton = driver.findElement(By.xpath("//*[@id=\"instanda-site-layout\"]/div[2]/div[2]/div/div/div/div/form/div[2]/input[2]"));
        clearviewendorsementbutton.click();
    }

    private void fillEndorseClause() {
        WebElement EndorseClauseContinueButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"instanda-site-layout\"]/div[2]/div[3]/div/form/div[2]/div/div/div[2]/button")));
        EndorseClauseContinueButton.click();
    }


    private void fillQuoteSummary(UWData d,  String filepath, int r) throws IOException, InterruptedException {
        WebElement quoteRefno = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/div[3]")));
        String quoteReference_number = quoteRefno.getText().trim();
        String quoteReference = "";
        try {
            quoteReference = quoteReference_number.replaceAll(".*Quote Reference:\\s*", "")
                    .replaceAll("\\s+.*", "")
                    .trim();
        } catch (Exception e) {
            quoteReference = "";
        }
        System.out.println("Extracted Quote Reference: " + quoteReference);
        Thread.sleep(3000); // for screenshot timing
        ExcelUtils.SetCellData(filepath, "UW Details", r, 16, quoteReference);
//        ExcelUtils.addScreenshotToWord(driver, doc,
//                 " Quote Reference Number : " + quoteReference, r);
       WebElement QuoteSummaryContinueButton = driver.findElement(By.id("continueButton"));
       QuoteSummaryContinueButton.click();
    }

    private void fillPaymentDetails() {
        WebElement PaymentPageContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[3]/div/div/div[2]/button[2]"));
        PaymentPageContinueButton.click();
    }
//    private String extractPolicyNumber() {
//        WebElement policyno = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/div[2]")));
//        String policynumber = policyno.getText();
//        System.out.println("Policy Created: " + policynumber);
//        return policynumber;
//    }
private String extractPolicyNumber(UWData d, XWPFDocument doc, String filepath, int r) throws InterruptedException, IOException {
    WebElement policyno = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/div[2]")));
    String policy_number = policyno.getText().trim();
    WebElement quoteRefno = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/div[3]")));
    String quoteReference_number = quoteRefno.getText().trim();
    // System.out.println("Policy Number for " + d.insuredName + " : " + policy_number);
    Thread.sleep(3000);
    // Capture Screenshot into Word (your utility method signature assumed)
  //  ExcelUtils.addScreenshotToWord(driver, doc, "Policy number for " + d.insuredName + " is: " + policy_number, r);
    // -----------------------------------------------
    // EXTRACT ONLY POLICY NUMBER → e.g. BX30089662
    // -----------------------------------------------
    String policyNumber = "";
    try {
        policyNumber = policy_number.replaceAll(".*Policy No\\.?:\\s*", "")
                .replaceAll("\\s+.*", "")
                .trim();
    } catch (Exception e) {
        policyNumber = "";
    }
    System.out.println("Extracted Policy Number: " + policyNumber);
    // -----------------------------------------------
    // EXTRACT ONLY QUOTE REFERENCE → e.g. E4V3J8
    // -----------------------------------------------
    String quoteReference = "";
    try {
        quoteReference = quoteReference_number.replaceAll(".*Quote Reference:\\s*", "")
                .replaceAll("\\s+.*", "")
                .trim();
    } catch (Exception e) {
        quoteReference = "";
    }
    System.out.println("Extracted Quote Reference: " + quoteReference);
    Thread.sleep(3000); // for screenshot timing
    // Screenshot into Word
    ExcelUtils.addScreenshotToWord(driver, doc,
            "Policy number for " + d.insuredName + " is: " + policyNumber + " | Quote Reference Number: " + quoteReference, r);

    //  ExcelUtils.SetCellData(filepath, "UW Details", r, 16, quoteReference);
    ExcelUtils.SetCellData(filepath, "UW Details", r, 17, policyNumber);

    if (!policyNumber.isEmpty()) {
        ExcelUtils.SetCellData(filepath, "UW Details", r, 18, "Success");
    } else {
        ExcelUtils.SetCellData(filepath, "UW Details", r, 18, "Failed");
        ExcelUtils.FillCellRed(filepath, "UW Details", r, 18);
    }
    return policyNumber;
}
}

