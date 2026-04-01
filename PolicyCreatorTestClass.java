package testPackSample;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import utils.ExcelUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.List;
import java.util.Random;

public class PolicyCreatorTestClass extends BaseClass {

    WebDriverWait wait;

    @Test
    public void CreateMultipleQuote() throws IOException, InterruptedException {

        String filepath = System.getProperty("user.dir") + "\\TestResources\\UWTestDataSampleAutomation.xlsx";
        //Getting Total Row Count in the Excel sheet
        int NofRows = ExcelUtils.getRowCount(filepath, "UW Datas");

        //Getting Total Col Counts in the Excel sheet
        int NofCols = ExcelUtils.getCellCount(filepath, "UW Datas", NofRows);
        wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        for (int r = 1; r <= 3; r++) {
            //Creating new word doc for Capturing evidence
            currentRow = r;   //  now r exists
            doc = new XWPFDocument();   //  create a fresh doc for each row
            //out = new FileOutputStream(System.getProperty("user.dir") + "\\TestResources\\TestEvidence.docx");
            out = new FileOutputStream(System.getProperty("user.dir") + "\\TestResources\\Evidence_Row_" + r + ".docx");     //Give each row its OWN evidence file

            UWData data = readExcelRow(filepath, r);
            driver.get("https://ops.digital-trading.int.hub.allianz.co.uk/Public/PreQuoteQuestions?PackageId=7&pageNumber=1");

            selectBrokerAccount(data);
            fillInsuredDetails(data);
            fillCoverDates(data);
            fillDeclarations(data);
            fillClaimsHistory(data);
            fillMaterialDamage(data);
            fillAdditionalCovers(data);
            // fillUnderwriterEditorsAndReferral();
            fillUnderwriterEditorsAndReferralIfPresent();
            fillEndorseClause();
            fillQuoteSummary(data, filepath,r,doc);
            fillPaymentDetails(data);
            extractPolicyNumber(data, doc, filepath, r);

            System.out.println("Quote created using row: " + r);
        }
    }
    // =============================================================
    // MODEL CLASS TO HOLD ONE ROW OF EXCEL DATA
    // =============================================================
    class UWData {
        String accountType, brokerName, contactName, contactPhone, contactEmail;
        String tradingStatus, companyReg, insuredName, tradingNameYN, tradingNameValue,subsidiaryCompanies,subsidiaryCompaniesCount,subsidiaryCompaniesName, postcode;
        String doesInsuredHaveSecondaryBusinessActivity,InsuredHaveSecondaryBusinessActivityName,estimatedSecondaryPercentageTurnover,turnoverPercentagePrimary;
        String establishedDate, turnover, businessActivity, coverStartDate,policyEndDateOptionYN,coverEndDate,reasonForNonStandardPolicyTenure,CountyCourtJudgementsDeclaration,CCJCount,CCJTotalMonetaryAmount,CCJRecentDate;
        String HMRevenueDeclarationYN,HMRevenueCountIfYes,LitigationDeclarationYN,LitigationCountIfYes;
        String claimHistoryInPast5YrsYN,claimCauseIfYes,claimOccurrenceDateIfYes,totalMonetaryAmountIfYes,selectDayOneUpliftPercentage;
        String totalDeclaredValueOfInstalledComputer, singleItemLossLimitForInstalledComputer,doestheInsuredRequirePortableComputerEquipmentYN,totaldeclaredValuePortableComputerEquipment,singleItemLossLimitForPortableComputerEquipment;
        String PortableComputerExcessTheft,PortableComputerExcessOtherClaims,PleaseConfirmYouWantToHaveShortTermPolicy,RenewalBehaviour;
        String requiredSumInsuredBreakdownBusinessInterruption,DoesInsuredRequireTerrorismCoverYN,ExcessPeriod,DoesInsuredHaveMaintenanceAgreementForComputerandAuxiliaryEquipment;
        String computerMediaAdditionalCoverYN,IndemnityPeriodBreakdownBusinessInterruption,WhatTypeBreakdownBusinessInterruptionCoverisRequired,DoesInsuredRequireBreakdownBusinessInterruptionYN,sumInsuredRrespectOfMaliciousCodeOrAttackCover,sumInsuredForComputerMediaAdditionalCover,doesInsuredRequireAdditionalExpenditureCoverYN,requiredSumInsuredExpenditureCover,requiredIndemnityPeriodAdditionalExpenditureCover,eRisksCoverYN,sumInsuredRespectOfSeekDestroyAndPreventCover;
    }

    // =============================================================
    // READ ONE ROW FROM EXCEL
    // =============================================================
    private UWData readExcelRow(String filepath, int r) throws IOException {
        UWData d = new UWData();

        d.accountType = ExcelUtils.getCellData(filepath, "UW Datas", r, 1);
        d.brokerName = ExcelUtils.getCellData(filepath, "UW Datas", r, 2);
        d.contactName = ExcelUtils.getCellData(filepath, "UW Datas", r, 3);
        d.contactPhone = ExcelUtils.getCellData(filepath, "UW Datas", r, 4);
        d.contactEmail = ExcelUtils.getCellData(filepath, "UW Datas", r, 5);
        d.tradingStatus = ExcelUtils.getCellData(filepath, "UW Datas", r, 6);
        d.companyReg = ExcelUtils.getCellData(filepath, "UW Datas", r, 7);
        d.insuredName = ExcelUtils.getCellData(filepath, "UW Datas", r, 8);

        d.tradingNameYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 9);
        d.tradingNameValue = ExcelUtils.getCellData(filepath, "UW Datas", r, 10);
        d.subsidiaryCompanies= ExcelUtils.getCellData(filepath, "UW Datas", r, 11);
        // d.subsidiaryCompaniesCount= ExcelUtils.getCellData(filepath, "UW Datas", r, 12);
        d.subsidiaryCompaniesName= ExcelUtils.getCellData(filepath, "UW Datas", r, 12);


        d.postcode = ExcelUtils.getCellData(filepath, "UW Datas", r, 13);
        d.establishedDate = ExcelUtils.getCellData(filepath, "UW Datas", r, 14);
        d.turnover = ExcelUtils.getCellData(filepath, "UW Datas", r, 15);
        d.businessActivity = ExcelUtils.getCellData(filepath, "UW Datas", r, 16);

        d.turnoverPercentagePrimary = ExcelUtils.getCellData(filepath, "UW Datas", r, 17);
        d.doesInsuredHaveSecondaryBusinessActivity = ExcelUtils.getCellData(filepath, "UW Datas", r, 18);
        d.InsuredHaveSecondaryBusinessActivityName = ExcelUtils.getCellData(filepath, "UW Datas", r, 19);
        d.estimatedSecondaryPercentageTurnover = ExcelUtils.getCellData(filepath, "UW Datas", r, 20);

        d.coverStartDate = ExcelUtils.getCellData(filepath, "UW Datas", r, 21);
        d.policyEndDateOptionYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 22);
        d.coverEndDate = ExcelUtils.getCellData(filepath, "UW Datas", r, 23);
        d.reasonForNonStandardPolicyTenure = ExcelUtils.getCellData(filepath, "UW Datas", r, 24);
        d.CountyCourtJudgementsDeclaration = ExcelUtils.getCellData(filepath, "UW Datas", r, 25);
        d.CCJCount = ExcelUtils.getCellData(filepath, "UW Datas", r, 26);
        d.CCJTotalMonetaryAmount = ExcelUtils.getCellData(filepath, "UW Datas", r, 27);
        d.CCJRecentDate = ExcelUtils.getCellData(filepath, "UW Datas", r, 28);

        d.HMRevenueDeclarationYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 29);
        d.HMRevenueCountIfYes = ExcelUtils.getCellData(filepath, "UW Datas", r, 30);
        d.LitigationDeclarationYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 31);
        d.LitigationCountIfYes = ExcelUtils.getCellData(filepath, "UW Datas", r, 32);

        d.claimHistoryInPast5YrsYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 33);
        d.claimCauseIfYes = ExcelUtils.getCellData(filepath, "UW Datas", r, 34);
        d.claimOccurrenceDateIfYes = ExcelUtils.getCellData(filepath, "UW Datas", r, 35);
        d.totalMonetaryAmountIfYes = ExcelUtils.getCellData(filepath, "UW Datas", r, 36);
        d.selectDayOneUpliftPercentage = ExcelUtils.getCellData(filepath, "UW Datas", r, 37);
        d.totalDeclaredValueOfInstalledComputer = ExcelUtils.getCellData(filepath, "UW Datas", r, 38);
        d.singleItemLossLimitForInstalledComputer = ExcelUtils.getCellData(filepath, "UW Datas", r, 39);
        d.doestheInsuredRequirePortableComputerEquipmentYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 40);
        d.totaldeclaredValuePortableComputerEquipment = ExcelUtils.getCellData(filepath, "UW Datas", r, 41);
        d.singleItemLossLimitForPortableComputerEquipment = ExcelUtils.getCellData(filepath, "UW Datas", r, 42);
        d.PortableComputerExcessTheft = ExcelUtils.getCellData(filepath, "UW Datas", r, 43);
        d.PortableComputerExcessOtherClaims = ExcelUtils.getCellData(filepath, "UW Datas", r, 44);

        d.computerMediaAdditionalCoverYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 45);
        d.sumInsuredForComputerMediaAdditionalCover = ExcelUtils.getCellData(filepath, "UW Datas", r, 46);
        d.doesInsuredRequireAdditionalExpenditureCoverYN= ExcelUtils.getCellData(filepath, "UW Datas", r, 47);
        d.requiredSumInsuredExpenditureCover = ExcelUtils.getCellData(filepath, "UW Datas", r, 48);
        d.requiredIndemnityPeriodAdditionalExpenditureCover = ExcelUtils.getCellData(filepath, "UW Datas", r, 49);
        d.eRisksCoverYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 50);
        d.sumInsuredRespectOfSeekDestroyAndPreventCover = ExcelUtils.getCellData(filepath, "UW Datas", r, 51);
        d.sumInsuredRrespectOfMaliciousCodeOrAttackCover = ExcelUtils.getCellData(filepath, "UW Datas", r, 52);
        d.DoesInsuredRequireBreakdownBusinessInterruptionYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 53);
        d.WhatTypeBreakdownBusinessInterruptionCoverisRequired = ExcelUtils.getCellData(filepath, "UW Datas", r, 54);
        d.IndemnityPeriodBreakdownBusinessInterruption = ExcelUtils.getCellData(filepath, "UW Datas", r, 55);
        d.requiredSumInsuredBreakdownBusinessInterruption = ExcelUtils.getCellData(filepath, "UW Datas", r, 56);
        d.DoesInsuredHaveMaintenanceAgreementForComputerandAuxiliaryEquipment = ExcelUtils.getCellData(filepath, "UW Datas", r, 57);
        d.ExcessPeriod = ExcelUtils.getCellData(filepath, "UW Datas", r, 58);
        d.DoesInsuredRequireTerrorismCoverYN = ExcelUtils.getCellData(filepath, "UW Datas", r, 59);
        // d.PleaseConfirmYouWantToHaveShortTermPolicy = ExcelUtils.getCellData(filepath, "UW Datas", r, 60);
        d.RenewalBehaviour = ExcelUtils.getCellData(filepath, "UW Datas", r, 60);

        return d;
    }


    private void jsClick(WebElement ele) {
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", ele);
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", ele);
    }
    // --- helpers ---
    private boolean isElementClickable(WebElement el) {
        try {
            return el.isDisplayed() && el.isEnabled();
        } catch (Exception e) {
            return false;
        }
    }

    private void hideSecondaryIfOpen(WebDriver driver) {
        // If your UI keeps the secondary block open due to previous state,
        // add steps to unselect/clear it here (collapse accordion, clear inputs, etc.).
    }
    private boolean getRandomBoolean() {
        return new Random().nextBoolean();
    }

    private void selectBrokerAccount(UWData d) throws InterruptedException, IOException {
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
        BrokerAccountName.click(); // Final click to close dropdown

        WebElement ContactName = driver.findElement(By.id("InsuranceAdviserName"));
        ContactName.sendKeys(d.contactName);
        WebElement ContactPhoneNumber = driver.findElement(By.id("InsuranceAdviserPhone"));
        ContactPhoneNumber.sendKeys(d.contactPhone);
        WebElement ContactEmailAddress = driver.findElement(By.id("InsuranceAdviserEmail"));
        ContactEmailAddress.sendKeys(d.contactEmail);

        captureStepScreenshot("SelectBrokerAccount Page ");

        WebElement AccountSelectionContinueBtn = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[2]/button[2]"));
        AccountSelectionContinueBtn.click();

    }



    private void fillInsuredDetails(UWData d) throws InterruptedException, IOException {

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/h1[2]")));
        WebElement LegaltradingStatus = driver.findElement(By.id("LegalTradingStatus"));
        Select tradingoption = new Select(LegaltradingStatus);

        String cleanName = d.tradingStatus                   // Clean Excel string (critical fix)
                .trim()
                .replace("\u00A0", "")      // remove non-breaking spaces
                .replaceAll("\\s+", " ");  // normalize multi-spaces
        System.out.println("Legal Trading status from excel: [" + cleanName + "]");
        tradingoption.selectByVisibleText(cleanName);

// Dynamic IDs according to the option selected from excel
        String insuredNameNumId = "";
        String companyRegNumId = "";
        String tradingNameRadioPrefixNo, tradingNameRadioPrefixYes = "";
        String subsidiaryRadioPrefixNo, subsidiaryRadioPrefixYes, SubsidiaryCompanyCount, nameOfSubsidiaryCompany, addMoreSubsidiaryCompanyAsNo = "";
        String tradingNameTextId = "";

        if (cleanName.equalsIgnoreCase("Private Limited")) {
            companyRegNumId = "StatusPrivateLimitedCompanyRegNumber";
            insuredNameNumId = "StatusPrivateLimitedInsuredName";
            tradingNameRadioPrefixNo = "StatusPrivateLimitedTradingNameYNNo";
            tradingNameRadioPrefixYes = "StatusPrivateLimitedTradingNameYNYes";
            tradingNameTextId = "StatusPrivateLimitedTradingName";
            subsidiaryRadioPrefixNo = "StatusPrivateLimitedSubsidiaryCoveredNo";
            subsidiaryRadioPrefixYes = "StatusPrivateLimitedSubsidiaryCoveredYes";
            SubsidiaryCompanyCount = "StatusPrivateLimitedNumSubsidiaries";
            nameOfSubsidiaryCompany = "StatusPrivateLimitedSubsidiaryName";
            addMoreSubsidiaryCompanyAsNo = "StatusPrivateLimitedAddSubsidiaryNo";
        } else if (cleanName.equalsIgnoreCase("Public Limited")) {
            companyRegNumId = "StatusPublicLimitedCompanyRegNumber";
            insuredNameNumId = "StatusPublicLimitedInsuredName";
            tradingNameRadioPrefixNo = "StatusPublicLimitedTradingNameYNNo";
            tradingNameRadioPrefixYes = "StatusPublicLimitedTradingNameYNYes";
            tradingNameTextId = "StatusPublicLimitedTradingName";
            subsidiaryRadioPrefixNo = "StatusPublicLimitedSubsidiaryCoveredNo";
            subsidiaryRadioPrefixYes = "StatusPublicLimitedSubsidiaryCoveredYes";
            SubsidiaryCompanyCount = "StatusPublicLimitedNumSubsidiaries";
            nameOfSubsidiaryCompany = "StatusPublicLimitedSubsidiaryName";
            addMoreSubsidiaryCompanyAsNo = "StatusPublicLimitedAddSubsidiaryNo";
        } else if (cleanName.equalsIgnoreCase("Private Unlimited")) {
            companyRegNumId = "StatusPrivateUnlimitedCompanyRegNumber";
            insuredNameNumId = "StatusPrivateUnlimitedInsuredName";
            tradingNameRadioPrefixNo = "StatusPrivateUnlimitedTradingNameYNNo";
            tradingNameRadioPrefixYes = "StatusPrivateUnlimitedTradingNameYNYes";
            tradingNameTextId = "StatusPrivateUnlimitedTradingName";
            subsidiaryRadioPrefixNo = "StatusPrivateUnlimitedSubsidiaryCoveredNo";
            subsidiaryRadioPrefixYes = "StatusPrivateUnlimitedSubsidiaryCoveredYes";
            SubsidiaryCompanyCount = "StatusPrivateUnlimitedNumSubsidiaries";
            nameOfSubsidiaryCompany = "StatusPrivateUnlimitedSubsidiaryName";
            addMoreSubsidiaryCompanyAsNo = "StatusPrivateUnlimitedAddSubsidiaryNo";
        } else if (cleanName.equalsIgnoreCase("Charity")) {
            companyRegNumId = "StatusCharityCompanyRegNumber";
            insuredNameNumId = "StatusCharityInsuredName";
            tradingNameRadioPrefixNo = "StatusCharityTradingNameYNNo";
            tradingNameRadioPrefixYes = "StatusCharityTradingNameYNYes";
            tradingNameTextId = "StatusCharityTradingName";
            subsidiaryRadioPrefixNo = "StatusCharitySubsidiaryCoveredNo";
            subsidiaryRadioPrefixYes = "StatusCharitySubsidiaryCoveredYes";
            SubsidiaryCompanyCount = "StatusCharityNumSubsidiaries";
            nameOfSubsidiaryCompany = "StatusCharitySubsidiaryName";
            addMoreSubsidiaryCompanyAsNo = "StatusCharityAddSubsidiaryNo";
        } else if (cleanName.equalsIgnoreCase("Trust")) {
            companyRegNumId = "StatusTrustCompanyRegNumber";
            insuredNameNumId = "StatusTrustInsuredName";
            tradingNameRadioPrefixNo = "StatusTrustTradingNameYNNo";
            tradingNameRadioPrefixYes = "StatusTrustTradingNameYNYes";
            tradingNameTextId = "StatusTrustTradingName";
            subsidiaryRadioPrefixNo = "StatusTrustSubsidiaryCoveredNo";
            subsidiaryRadioPrefixYes = "StatusTrustSubsidiaryCoveredYes";
            SubsidiaryCompanyCount = "StatusTrustNumSubsidiaries";
            nameOfSubsidiaryCompany = "StatusTrustSubsidiaryName";
            addMoreSubsidiaryCompanyAsNo = "StatusTrustAddSubsidiaryNo";
        } else if (cleanName.equalsIgnoreCase("Limited Liability Partnership")) {
            companyRegNumId = "StatusLimitedLiabilityPartnershipCompanyRegNumber";
            insuredNameNumId = "StatusLimitedLiabilityPartnershipInsuredName";
            tradingNameRadioPrefixNo = "StatusLimitedLiabilityPartnershipTradingNameYNNo";
            tradingNameRadioPrefixYes = "StatusLimitedLiabilityPartnershipTradingNameYNYes";
            tradingNameTextId = "StatusLimitedLiabilityPartnershipTradingName";
            subsidiaryRadioPrefixNo = "StatusLimitedLiabilityPartnershipSubsidiaryCoveredNo";
            subsidiaryRadioPrefixYes = "StatusLimitedLiabilityPartnershipSubsidiaryCoveredYes";
            SubsidiaryCompanyCount = "StatusLimitedLiabilityPartnershipNumSubsidiaries";
            nameOfSubsidiaryCompany = "StatusLimitedLiabilityPartnershipSubsidiaryName";
            addMoreSubsidiaryCompanyAsNo = "StatusLimitedLiabilityPartnershipAddSubsidiaryNo";
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
        //--------------------------------------------------------------
        //             TRADING NAME:     YES / NO FROM EXCEL
        if (d.tradingNameYN.equalsIgnoreCase("YES")) {
            WebElement yesButton = driver.findElement(By.id(tradingNameRadioPrefixYes));
            jsClick(yesButton);
            WebElement text = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(tradingNameTextId)));
            text.sendKeys(d.tradingNameValue);
        } else {
            WebElement noButton = driver.findElement(By.id(tradingNameRadioPrefixNo));
            jsClick(noButton);
        }
//--------------------------------------------------------------
////             Subsidiary : yes/No from excel
//--------------------------------------------------------------
        if (d.subsidiaryCompanies.equalsIgnoreCase("YES")) {
            By subsidiaryYes = By.xpath("//label[.//input[starts-with(@id,'" + subsidiaryRadioPrefixYes + "') and @value='Yes']]");
            WebElement subsidiaryLabel = wait.until(ExpectedConditions.presenceOfElementLocated(subsidiaryYes));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", subsidiaryLabel);

            WebElement subsidiaryCompaniesCount = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(SubsidiaryCompanyCount)));
            Select countSelect = new Select(subsidiaryCompaniesCount); // here not picking Subsidiarycount from excel,intead value hardcoded
            countSelect.selectByValue("1");  //or    countSelect.selectByVisibleText("1");
            WebElement subCompanyName = driver.findElement(By.id(nameOfSubsidiaryCompany));
            subCompanyName.sendKeys(d.subsidiaryCompaniesName);
            WebElement addsubCompanyNo = driver.findElement(By.id(addMoreSubsidiaryCompanyAsNo));
            jsClick(addsubCompanyNo);
        } else {
            By subsidiaryNo = By.xpath("//label[.//input[starts-with(@id,'" + subsidiaryRadioPrefixNo + "') and @value='No']]");
            WebElement subsidiaryLabelNo = wait.until(ExpectedConditions.presenceOfElementLocated(subsidiaryNo));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", subsidiaryLabelNo);
        }

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
            if (option.getText().contains(d.businessActivity)) {
                JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;
                jsExecutor.executeScript("arguments[0].click();", option);
                System.out.println("BusinessActivity selected :" + d.businessActivity );
                WebElement desiredoption = driver.findElement(By.cssSelector(".tt-suggestion.tt-selectable,.tt-highlight"));
                desiredoption.click();
                break;
            }
        }

        String secondaryBusinessFlag = d.doesInsuredHaveSecondaryBusinessActivity;
        String primaryPctStr = d.turnoverPercentagePrimary;  // whatever column holds it
        String secondaryPctStr = d.estimatedSecondaryPercentageTurnover;
        //----------------------------concept-----------------------------------------------
//        The system requires total estimated turnover = 100%
//                If there is only 1 business activity, then that one must be 100%
//                If secondary business = NO, then there should be no second percentage and no validation error
//        If secondary business = YES, then both percentages must add to 100% (Example: 80 + 20, 70 + 30, 100 + 0)
        //---------------------------------------------------------------------------------------------

// --- Parse and guard against bad input ---
        int primaryPct = 0;
        try {
            primaryPct = Integer.parseInt(primaryPctStr.trim());
        } catch (Exception e) {
            throw new RuntimeException("Primary % is not a valid integer: " + primaryPctStr);
        }
        if (primaryPct < 0 || primaryPct > 100) {
            throw new RuntimeException("Primary % must be between 0 and 100. Got: " + primaryPct);
        }

// --- Locate elements (use your existing locators where applicable) ---
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement primaryPctInput = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.id("EstimatedPercentageTurnover")));
        primaryPctInput.clear();


        WebElement secondaryYesBtn = wait.until(ExpectedConditions.visibilityOfElementLocated(
                By.xpath("//*[@id=\"question1091\"]/div[1]/div[2]/div[1]/div[1]/label[1]")));
        WebElement secondaryNoBtn  = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"question1091\"]/div[1]/div[2]/div[1]/div[1]/label[2]")));

// --- Branch on Secondary flag ---
        if ("No".equalsIgnoreCase(secondaryBusinessFlag)) {
            // Force primary = 100 and skip secondary block entirely
            primaryPctInput.clear();
            primaryPctInput.sendKeys("100");

            // Make sure NO is selected (do not open secondary fields)
            if (isElementClickable(secondaryNoBtn)) secondaryNoBtn.click();

            // Optional: if the page automatically shows previous secondary fields, collapse them
            hideSecondaryIfOpen(driver);

        } else {
            // User said YES → fill both so that total = 100
            if (isElementClickable(secondaryYesBtn)) secondaryYesBtn.click();
            // Primary: use the Excel value (but if it is blank, default to 100)
            primaryPctInput.clear();
            primaryPctInput.sendKeys(String.valueOf(primaryPct));
            // Secondary: from Excel if provided, else compute 100 - primary
            int secondaryPct;
            if (secondaryPctStr != null && !secondaryPctStr.trim().isEmpty()) {
                try {
                    secondaryPct = Integer.parseInt(secondaryPctStr.trim());
                } catch (Exception e) {
                    throw new RuntimeException("Secondary % is not a valid integer: " + secondaryPctStr);
                }
            } else {
                secondaryPct = 100 - primaryPct;
            }
            // Guard rails
            if (secondaryPct < 0 || secondaryPct > 100) {
                throw new RuntimeException("Secondary % must be between 0 and 100. Got: " + secondaryPct);
            }
            if (primaryPct + secondaryPct != 100) {
                throw new RuntimeException("Primary + Secondary must total 100. Got: " + primaryPct + " + " + secondaryPct);
            }
            // Type Secondary %
            WebElement secondaryPctInput = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.xpath("//*[@id=\"EstimatedPercentageTurnover1\"]") // <-- replace with your actual id
            ));
            secondaryPctInput.clear();
            secondaryPctInput.sendKeys(String.valueOf(secondaryPct));

            //---------------------------------------------
            WebElement SecondBusinessActivity = driver.findElement(By.xpath("//*[@id=\"Instanda_TradesUndertaken\"]"));
            SecondBusinessActivity.sendKeys(d.InsuredHaveSecondaryBusinessActivityName);
            List<WebElement> secOptions = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("Instanda_TradesUndertaken_listbox")));
            for (WebElement opt : secOptions) {
                if (opt.getText().trim().equalsIgnoreCase(d.InsuredHaveSecondaryBusinessActivityName.trim())) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", opt);
                    System.out.println("Secondary Business Selected: " + opt.getText());
                    break;
                }
            }

            WebElement doesInsuredHaveThirdBusinessActivity = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.xpath("//*[@id=\"question8028\"]/div[1]/div[2]/div[1]/div[1]/label[2]"))); // <-- selecting No
            doesInsuredHaveThirdBusinessActivity.click();
        }

        captureStepScreenshot("Insured Details Page ");
        // Finally, continue
        WebElement InsuredContinueButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]")));
        InsuredContinueButton.click();

    }

    private void fillCoverDates(UWData d) throws IOException {
        WebElement PolicyInceptionDate = driver.findElement(By.id("PolicyInceptionDate"));
        PolicyInceptionDate.clear();
        PolicyInceptionDate.sendKeys(d.coverStartDate);
        //----------------------------------------------------------
        //             policyEndDate :     YES / NO FROM EXCEL
        if (d.policyEndDateOptionYN.equalsIgnoreCase("YES")) {
            WebElement PolicyEndDateYes = driver.findElement(By.xpath("//*[@id=\"question1082\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement PolicyEndDateOptionYes = PolicyEndDateYes.findElement(By.xpath("//*[@id=\"question1082\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            jsClick(PolicyEndDateOptionYes);          // OR PolicyEndDateOptionYes.click();
            WebElement CoverEndDate = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"PolicyLapseDateCustom\"]")));
            CoverEndDate.sendKeys(d.coverEndDate);
            WebElement reasonNonStandardPolicySelected = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ReasonNonStandardTerm")));
            Select drp = new Select(reasonNonStandardPolicySelected);         // dropdown pick from excel
            drp.selectByVisibleText(d.reasonForNonStandardPolicyTenure.trim());
        } else {         //No from Excel
            WebElement PolicyEndDateNo = driver.findElement(By.xpath("//*[@id=\"question1082\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement PolicyEndDateOptionNo = PolicyEndDateNo.findElement(By.xpath("//*[@id=\"question1082\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            jsClick(PolicyEndDateOptionNo);
        }

        captureStepScreenshot("CoverDates Page ");
        WebElement CoverDatesContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        CoverDatesContinueButton.click();
    }

    private void fillDeclarations(UWData d) throws IOException {
        WebElement CriminalConvictions = driver.findElement(By.xpath("//*[@id=\"question1722\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement CCOptionNo = CriminalConvictions.findElement(By.xpath("//*[@id=\"question1722\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        CCOptionNo.click();
        WebElement BankandLiquidations = driver.findElement(By.xpath("//*[@id=\"question1727\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement BankandLiquidationsOption = BankandLiquidations.findElement(By.xpath("//*[@id=\"question1727\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        BankandLiquidationsOption.click();
        WebElement InsuranceVoid = driver.findElement(By.xpath("//*[@id=\"question1731\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement InsuranceVoidOption = InsuranceVoid.findElement(By.xpath("//*[@id=\"question1731\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        InsuranceVoidOption.click();

        //             CountyCourtJudgementsDeclaration :     YES / NO FROM EXCEL
        if (d.CountyCourtJudgementsDeclaration.equalsIgnoreCase("YES")) {
            WebElement CCJs = driver.findElement(By.xpath("//*[@id=\"question1733\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement CCJsOptionYes = CCJs.findElement(By.xpath("//*[@id=\"question1733\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            CCJsOptionYes.click();
            WebElement CCJsinLast5years = CCJs.findElement(By.xpath("//*[@id=\"CCJCountLast5Years\"]"));
            CCJsinLast5years.sendKeys(d.CCJCount);
            WebElement totalMonetaryAmountAssociatedWithCCJ = CCJs.findElement(By.xpath("//*[@id=\"CCJTotalAmountLast5Years\"]"));
            totalMonetaryAmountAssociatedWithCCJ.sendKeys(d.CCJTotalMonetaryAmount);
            WebElement dateOfMostRecentCCJ  = CCJs.findElement(By.xpath("//*[@id=\"CCJMostRecentDate\"]"));
            dateOfMostRecentCCJ.sendKeys(d.CCJRecentDate);
        } else {
            WebElement CCJs = driver.findElement(By.xpath("//*[@id=\"question1733\"]/div[1]/div[2]/div[1]/div[1]"));   //No from Excel
            WebElement CCJsOptionNo = CCJs.findElement(By.xpath("//*[@id=\"question1733\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            CCJsOptionNo.click();
        }

        WebElement Disqualifications = driver.findElement(By.xpath("//*[@id=\"question1746\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement DisqualificationOption = Disqualifications.findElement(By.xpath("//*[@id=\"question1746\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        DisqualificationOption.click();

        //             HMRevenueOptionDeclaration :     YES / NO FROM EXCEL
        if (d.HMRevenueDeclarationYN.equalsIgnoreCase("YES")) {
            WebElement HMRevenue = driver.findElement(By.xpath("//*[@id=\"question1756\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement HMRevenueOptionYes = HMRevenue.findElement(By.xpath("//*[@id=\"question1756\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            HMRevenueOptionYes.click();
            WebElement HMRevenueNumberOfOccurence = driver.findElement(By.xpath("//*[@id=\"HMRCRecoveryActionOccuranceLast5Years\"]"));
            HMRevenueNumberOfOccurence.sendKeys(d.HMRevenueCountIfYes);
        } else {
            WebElement HMRevenue = driver.findElement(By.xpath("//*[@id=\"question1756\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement HMRevenueOptionNo = HMRevenue.findElement(By.xpath("//*[@id=\"question1756\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            HMRevenueOptionNo.click();
        }

        WebElement HealthyandSafety = driver.findElement(By.xpath("//*[@id=\"question1758\"]/div[1]/div[2]/div[1]/div[1]"));
        WebElement HealthyandSafetyOption = HealthyandSafety.findElement(By.xpath("//*[@id=\"question1758\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
        HealthyandSafetyOption.click();

        //             LitigationDeclaration :     YES / NO FROM EXCEL
        if (d.LitigationDeclarationYN.equalsIgnoreCase("YES")) {
            WebElement Litigation = driver.findElement(By.xpath("//*[@id=\"question1765\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement LitigationOptionYes = Litigation.findElement(By.xpath("//*[@id=\"question1765\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));          //*[@id="LitigationDisputesLast5YearsYes"]
            LitigationOptionYes.click();
            WebElement LitigationDisputesNumOccuranceLast5Years = driver.findElement(By.xpath("//*[@id=\"LitigationDisputesOccuranceLast5Years\"]"));
            LitigationDisputesNumOccuranceLast5Years.sendKeys(d.LitigationCountIfYes);
        } else {
            WebElement Litigation = driver.findElement(By.xpath("//*[@id=\"question1765\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement LitigationOptionNo = Litigation.findElement(By.xpath("//*[@id=\"question1765\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            LitigationOptionNo.click();
        }
        captureStepScreenshot("Declarations Page ");
        WebElement DeclarationsContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        DeclarationsContinueButton.click();
    }
    private void fillClaimsHistory(UWData d) throws InterruptedException, IOException {

//             claimHistoryInPast5YrsYN :     YES / NO FROM EXCEL
        if (d.claimHistoryInPast5YrsYN.equalsIgnoreCase("YES")) {
            WebElement AddClaimsInThePast5Years = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"ClaimsInThePast5YearsaddButton\"]")));
            AddClaimsInThePast5Years.click();
            WebElement ClaimCauseDropdown = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div/form/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/select"));
            jsClick(ClaimCauseDropdown);
            Thread.sleep(500);
            // Create Select object for dropdown
            Select select = new Select(ClaimCauseDropdown);
            String cleanName = d.claimCauseIfYes                   // Clean Excel string (critical fix)
                    .trim()
                    .replace("\u00A0", "")      // remove non-breaking spaces
                    .replaceAll("\\s+", " ");  // normalize multi-spaces
            System.out.println("ClaimCause from excel: [" + cleanName + "]");
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
            ClaimCauseDropdown.click(); // Final click to close dropdown
            WebElement claimOccurrenceDate = driver.findElement(By.xpath("//*[@id=\"ClaimsInThePast5Years1_ClaimOccurrenceDate\"]"));
            claimOccurrenceDate.sendKeys(d.claimOccurrenceDateIfYes);
            WebElement totalMonetaryAmount = driver.findElement(By.xpath("//*[@id=\"ClaimsInThePast5Years1_ValueOfClaimsInLast5Years\"]"));
            totalMonetaryAmount.sendKeys(d.totalMonetaryAmountIfYes);
        }
        else {  //if selecting No from excel
            WebElement saveButton = scrollUntilVisible(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[2]"));
            saveButton.click();
        }
        Thread.sleep(1000);
        captureStepScreenshot("ClaimsHistory Page ");
        WebElement ClaimsHistoryContinueButton = wait.until(ExpectedConditions.elementToBeClickable((By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"))));
        jsClick(ClaimsHistoryContinueButton);
    }

    private void fillMaterialDamage(UWData d) throws InterruptedException, IOException {
        WebElement selectDayOneUpliftPercentage = driver.findElement(By.xpath("//*[@id=\"DayOneUplift\"]"));
        jsClick(selectDayOneUpliftPercentage);
        Thread.sleep(500);
        // Create Select object for dropdown
        Select select = new Select(selectDayOneUpliftPercentage);
        String cleanName = d.selectDayOneUpliftPercentage                   // Clean Excel string (critical fix)
                .trim()
                .replace("\u00A0", "")      // remove non-breaking spaces
                .replaceAll("\\s+", " ");  // normalize multi-spaces
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
        //Premises 1------------------
        WebElement yesRadioButton = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@name='PremisesIsSameAsClientAddress__11__1' and @value='Yes']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", yesRadioButton);
        Thread.sleep(1000);
        //---------------------------------------------------------------

// Condition handling : Material Damage Page ,If 2 Business Activity exist ,then dropdown comes and we select 2nd one here, if one business activity exist then no dropdown (else condition)
        if (d.doesInsuredHaveSecondaryBusinessActivity.equalsIgnoreCase("Yes")) { ////(Column S)
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            WebElement secondaryPremisesBusinessActivityDropdown = wait.until(
                    ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"Premises1_Instanda_ActivitiesAtThisLocation\"]")));
            Select selectSecondary = new Select(secondaryPremisesBusinessActivityDropdown);
            selectSecondary.selectByVisibleText(d.InsuredHaveSecondaryBusinessActivityName);  //(Column T)
            System.out.println("Material Damage Page ,Secondary Business Selected: " + d.InsuredHaveSecondaryBusinessActivityName);
        } else {
            // // Select primary business activity (Column Q) //No dropdown appears
            WebElement premises1BusinessActivity = driver.findElement(By.xpath("//*[@id=\"Premises1_Instanda_ActivitiesAtThisLocation\"]"));
            Select selectPrimary = new Select(premises1BusinessActivity);
            selectPrimary.selectByVisibleText(d.businessActivity);
            System.out.println("Material Damage Page ,Primary Business Activity only exist: " + d.InsuredHaveSecondaryBusinessActivityName);
        }

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

        //             portableComputer :     YES / NO FROM EXCEL
        if (d.doestheInsuredRequirePortableComputerEquipmentYN.equalsIgnoreCase("YES")) {
            WebElement PortableComputer = driver.findElement(By.xpath("//*[@id=\"question2100\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement PortableComputerOptionYes = PortableComputer.findElement(By.xpath("//*[@id=\"question2100\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            PortableComputerOptionYes.click();

            WebElement totaldeclaredValuePPortableComputerEquipment = driver.findElement(By.xpath("//*[@id=\"PortableComputerEquipmentDeclaredValue\"]"));
            totaldeclaredValuePPortableComputerEquipment.sendKeys(d.totaldeclaredValuePortableComputerEquipment);
            WebElement PortableComputerEquipmentSigleItemLossLimit = driver.findElement(By.xpath("//*[@id=\"PortableComputerEquipmentLossLimit\"]"));
            PortableComputerEquipmentSigleItemLossLimit.sendKeys(d.singleItemLossLimitForPortableComputerEquipment);


            WebElement portableComputerExcessTheftMaliciousDamage = driver.findElement(By.xpath("//*[@id=\"PortableTheftExcessInitial\"]"));
            portableComputerExcessTheftMaliciousDamage.sendKeys(d.PortableComputerExcessTheft);
            WebElement portableComputerExcessAllOtherClaims = driver.findElement(By.xpath("//*[@id=\"PortableAllOtherExcessInitial\"]"));
            portableComputerExcessAllOtherClaims.sendKeys(d.PortableComputerExcessOtherClaims);

        } else {  //Portable Computer No condition
            WebElement PortableComputer = driver.findElement(By.xpath("//*[@id=\"question2100\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement PortableComputerOptionNo = PortableComputer.findElement(By.xpath("//*[@id=\"question2100\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            PortableComputerOptionNo.click();
        }
        captureStepScreenshot("MaterialDamage Page ");
        WebElement MaterialDamageContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        MaterialDamageContinueButton.click();


    }

    private void fillAdditionalCovers(UWData d) throws IOException {
//--------------concept---------------------
        //• If E-Risk = YES ➜ Computer Media (AT) = YES AND Additional Exp (AV) = YES (mandatory)
        //• If E-Risk = NO ➜ No restriction (anything is fine)
        //IF (AY == YES)
        //   → AT must be YES
        //   → AV must be YES
        //ELSE
        //   → No dependency (just select values as per Excel)

// ====== MAIN CONDITION ======
        if (d.eRisksCoverYN.equalsIgnoreCase("Yes")) {
            // Force both dependencies to YES (Even if Excel says "NO", these three options are forced to YES when E-risk is selected as Yes.)
            WebElement ComputerMedia = driver.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement ComputerMediaOptionYes = ComputerMedia.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            ComputerMediaOptionYes.click();
            WebElement ComputerMediaSumInsured = driver.findElement(By.xpath("//*[@id=\"ComputerMediaSumInsured\"]"));
            ComputerMediaSumInsured.sendKeys(d.sumInsuredForComputerMediaAdditionalCover);
            WebElement AdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement AdditionalExpenditureOptionYes = AdditionalExpenditure.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            AdditionalExpenditureOptionYes.click();
            WebElement requiredSumInsuredForAdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"AdditionalExpenditureSumInsured\"]"));
            requiredSumInsuredForAdditionalExpenditure.sendKeys(d.requiredSumInsuredExpenditureCover);
            WebElement requiredIndemnityPeriodForAdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"AdditionalExpenditureIndemnityPeriodChoice\"]"));
            requiredIndemnityPeriodForAdditionalExpenditure.sendKeys(d.requiredIndemnityPeriodAdditionalExpenditureCover);

            WebElement ERisks = driver.findElement(By.xpath("//*[@id=\"question1214\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement ERisksOptionYes = ERisks.findElement(By.xpath("//*[@id=\"question1214\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            ERisksOptionYes.click();
            WebElement ERisksSeekAndDestroySumInsured = driver.findElement(By.xpath("//*[@id=\"ERisksSeekAndDestroySumInsured\"]"));
            ERisksSeekAndDestroySumInsured.sendKeys(d.sumInsuredRespectOfSeekDestroyAndPreventCover);
            WebElement ERisksVirusAndHackingSumInsured = driver.findElement(By.xpath("//*[@id=\"ERisksVirusAndHackingSumInsured\"]"));
            ERisksVirusAndHackingSumInsured.sendKeys(d.sumInsuredRrespectOfMaliciousCodeOrAttackCover);

        } else {
            if (d.computerMediaAdditionalCoverYN.equalsIgnoreCase("Yes")) {
                WebElement ComputerMedia = driver.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]"));
                WebElement ComputerMediaOptionYes = ComputerMedia.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
                ComputerMediaOptionYes.click();
                WebElement ComputerMediaSumInsured = driver.findElement(By.xpath("//*[@id=\"ComputerMediaSumInsured\"]"));
                ComputerMediaSumInsured.sendKeys(d.sumInsuredForComputerMediaAdditionalCover);
            } else {
                WebElement ComputerMedia = driver.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]"));
                WebElement ComputerMediaOptionNo = ComputerMedia.findElement(By.xpath("//*[@id=\"question1219\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
                ComputerMediaOptionNo.click();
            }
            if (d.doesInsuredRequireAdditionalExpenditureCoverYN.equalsIgnoreCase("Yes")) {
                WebElement AdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]"));
                WebElement AdditionalExpenditureOptionYes = AdditionalExpenditure.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
                AdditionalExpenditureOptionYes.click();
                WebElement requiredSumInsuredForAdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"AdditionalExpenditureSumInsured\"]"));
                requiredSumInsuredForAdditionalExpenditure.sendKeys(d.requiredSumInsuredExpenditureCover);
                WebElement requiredIndemnityPeriodForAdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"AdditionalExpenditureIndemnityPeriodChoice\"]"));
                requiredIndemnityPeriodForAdditionalExpenditure.sendKeys(d.requiredIndemnityPeriodAdditionalExpenditureCover);
            } else {
                WebElement AdditionalExpenditure = driver.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]"));
                WebElement AdditionalExpenditureOptionNo = AdditionalExpenditure.findElement(By.xpath("//*[@id=\"question1208\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
                AdditionalExpenditureOptionNo.click();
            }
            WebElement ERisks = driver.findElement(By.xpath("//*[@id=\"question1214\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement ERisksOptionNo = ERisks.findElement(By.xpath("//*[@id=\"question1214\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            ERisksOptionNo.click();
        }
        if (d.DoesInsuredRequireBreakdownBusinessInterruptionYN.equalsIgnoreCase("Yes")) {
            WebElement BreakdownBusiness = driver.findElement(By.xpath("//*[@id=\"question1195\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement BreakdownBusinessOptionYes = BreakdownBusiness.findElement(By.xpath("//*[@id=\"question1195\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
            BreakdownBusinessOptionYes.click();
            WebElement BreakdownBusinessInterruptionCoverBasis = driver.findElement(By.xpath("//*[@id=\"BreakdownBusinessInterruptionCoverBasis\"]"));
            BreakdownBusinessInterruptionCoverBasis.sendKeys(d.WhatTypeBreakdownBusinessInterruptionCoverisRequired);
            WebElement BreakdownBusinessInterruptionIndemnityPeriod = driver.findElement(By.xpath("//*[@id=\"BreakdownBusinessInterruptionIndemnityPeriod\"]"));
            BreakdownBusinessInterruptionIndemnityPeriod.sendKeys(d.IndemnityPeriodBreakdownBusinessInterruption);
            WebElement BreakdownBusinessInterruptionSumInsured = driver.findElement(By.xpath("//*[@id=\"BreakdownBusinessInterruptionSumInsured\"]"));
            BreakdownBusinessInterruptionSumInsured.sendKeys(d.requiredSumInsuredBreakdownBusinessInterruption);
            if (d.DoesInsuredHaveMaintenanceAgreementForComputerandAuxiliaryEquipment.equalsIgnoreCase("Yes")) {
                WebElement MaintenanceAgreementComputerandAuxiliaryEquipmentYes = driver.findElement(By.xpath("//*[@id=\"question2986\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
                MaintenanceAgreementComputerandAuxiliaryEquipmentYes.click();
                WebElement BIExcessPeriodWithAgreement = driver.findElement(By.xpath("//*[@id=\"BIExcessPeriodWithAgreement\"]"));
                BIExcessPeriodWithAgreement.sendKeys(d.ExcessPeriod);
            } else {
                WebElement MaintenanceAgreementComputerandAuxiliaryEquipmentNo = driver.findElement(By.xpath("//*[@id=\"question2986\"]/div[1]/div[2]/div[1]/div[1]/label[1]"));
                MaintenanceAgreementComputerandAuxiliaryEquipmentNo.click();
                WebElement BIExcessPeriodWithAgreement = driver.findElement(By.xpath("//*[@id=\"BIExcessPeriodWithAgreement\"]"));
                BIExcessPeriodWithAgreement.sendKeys(d.ExcessPeriod);
            }

        } else {
            WebElement BreakdownBusiness = driver.findElement(By.xpath("//*[@id=\"question1195\"]/div[1]/div[2]/div[1]/div[1]"));
            WebElement BreakdownBusinessOptionNo = BreakdownBusiness.findElement(By.xpath("//*[@id=\"question1195\"]/div[1]/div[2]/div[1]/div[1]/label[2]"));
            BreakdownBusinessOptionNo.click();
        }

        if (d.DoesInsuredRequireTerrorismCoverYN.equalsIgnoreCase("Yes")) {
            WebElement Terrorism = driver.findElement(By.xpath("//*[@id=\"question1217\"]/div[2]/div[1]/div/div[1]"));
            WebElement TerrorismOptionYes = Terrorism.findElement(By.xpath("//*[@id=\"question1217\"]/div[2]/div[1]/div/div[1]/label[1]"));
            TerrorismOptionYes.click();
        } else {
            WebElement Terrorism = driver.findElement(By.xpath("//*[@id=\"question1217\"]/div[2]/div[1]/div/div[1]"));
            WebElement TerrorismOptionNo = Terrorism.findElement(By.xpath("//*[@id=\"question1217\"]/div[2]/div[1]/div/div[1]/label[2]"));
            TerrorismOptionNo.click();
        }
        captureStepScreenshot("AdditionalCovers Page ");
        WebElement AdditonalPageContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[2]/div/div/div[3]/button[3]"));
        AdditonalPageContinueButton.click();
    }
    public boolean isReferralPagePresent() {
        try {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
            wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//input[@type='checkbox']")));
            return true;
        } catch (TimeoutException e) {
            return false;  // No referral page
        }
    }
    private void fillUnderwriterEditorsAndReferralIfPresent() throws InterruptedException, IOException {

        if (isReferralPagePresent()) {
            System.out.println("Referral detected → Handling referral page...");
            captureStepScreenshot("Referral detected → Handling referral page");
            fillUnderwriterEditorsAndReferral();
        } else {
            System.out.println("Referral NOT detected → Skipping to EndorseClause...");
        }
    }
    private void fillUnderwriterEditorsAndReferral() throws IOException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        //  Wait for ANY checkbox
        List<WebElement> checkboxes = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//input[@type='checkbox']")));
        //  Loop through all checkboxes & select them
        for (WebElement checkbox : checkboxes) {
            try {
                if (!checkbox.isSelected()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", checkbox);
                    wait.until(ExpectedConditions.elementToBeClickable(checkbox));
                    checkbox.click();
                }
            } catch (Exception e) {
                //  Fall back to JS click
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", checkbox);
            }
        }
        // Click CLEAR button
        WebElement clearButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"js-clearButton\"]")));
        clearButton.click();

        //OPTIONAL POP‑UP (handle only if it appears)
        try {
            WebElement confirm = new WebDriverWait(driver, Duration.ofSeconds(3))
                    .until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"clear-confirm\"]")));
            confirm.click();
        } catch (TimeoutException ignored) { }
        // Enter reason
        WebElement reasontextbox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("Reason")));
        reasontextbox.clear();
        reasontextbox.sendKeys("Cleared all referrals");

        captureStepScreenshot("UW EditorAndReferral Page ");
        // Click Continue / Clear View Endorsement button
        WebElement clearViewEndorsementButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"instanda-site-layout\"]/div[2]/div[2]/div/div/div/div/form/div[2]/input[2]")));
        clearViewEndorsementButton.click();
    }

    private void fillEndorseClause() throws InterruptedException, IOException {
        captureStepScreenshot("EndorseClause Page ");
        WebElement EndorseClauseContinueButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"instanda-site-layout\"]/div[2]/div[3]/div/form/div[2]/div/div/div[2]/button")));
        EndorseClauseContinueButton.click();
    }


    private void fillQuoteSummary(UWData d,  String filepath, int r,XWPFDocument doc) throws IOException, InterruptedException {
        Thread.sleep(500);
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
        ExcelUtils.SetCellData(filepath, "UW Datas", r, 61, quoteReference);
        captureStepScreenshot("QuoteSummary Page ");
        WebElement QuoteSummaryContinueButton = driver.findElement(By.id("continueButton"));
        QuoteSummaryContinueButton.click();
    }
    private void fillPaymentDetails(UWData d) throws IOException {
        WebElement RenewalBehaviour = driver.findElement(By.xpath("//*[@id=\"RenewalBehaviour\"]"));
        RenewalBehaviour.sendKeys(d.RenewalBehaviour);
        captureStepScreenshot("Payment Details Page ");
        WebElement PaymentPageContinueButton = driver.findElement(By.xpath("//*[@id=\"instandaquestions\"]/div[3]/div/div/div[2]/button[2]"));
        PaymentPageContinueButton.click();
    }

    private String extractPolicyNumber(UWData d, XWPFDocument doc, String filepath, int r) throws InterruptedException, IOException {
        WebElement policyno = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/div[2]")));
        String policy_number = policyno.getText().trim();
        WebElement quoteRefno = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"page-title\"]/div[1]/div[3]")));
        String quoteReference_number = quoteRefno.getText().trim();
        Thread.sleep(3000);
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
        ExcelUtils.addScreenshotToWord(driver, doc, "Extracted Policy Number : " + policyNumber + " & Quote Reference Num : " + quoteReference + "  For " +  d.insuredName, r);

        //  ExcelUtils.SetCellData(filepath, "UW Datas", r, 61, quoteReference);
        ExcelUtils.SetCellData(filepath, "UW Datas", r, 62, policyNumber);
        if (!policyNumber.isEmpty()) {
            ExcelUtils.SetCellData(filepath, "UW Datas", r, 63, "Pass");
            ExcelUtils.FillCellGreen(filepath, "UW Datas", r, 63);
        } else {
            ExcelUtils.SetCellData(filepath, "UW Datas", r, 64, "Fail");
            ExcelUtils.FillCellRed(filepath, "UW Datas", r, 64);
        }
        return policyNumber;
    }
}