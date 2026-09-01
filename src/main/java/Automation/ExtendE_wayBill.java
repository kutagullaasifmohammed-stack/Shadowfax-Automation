package Automation;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class ExtendE_wayBill {

    public static void main(String[] args) throws Exception {

        // ================= EXCEL INITIAL READING =================
        String path = ".\\TestData1\\AutoExtendFile.xlsx";
        FileInputStream fis = null;
        XSSFWorkbook workbook = null;

        try {
            fis = new FileInputStream(path);
            workbook = new XSSFWorkbook(fis);
            System.out.println("[INFO] Excel file loaded successfully.");
        } catch (Exception e) {
            System.err.println("CRITICAL ERROR: Make sure 'AutoExtendFile.xlsx' is CLOSED in Excel before running!");
            throw e;
        } finally {
            if (fis != null) fis.close();
        }

        XSSFSheet sheet = workbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter();

        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        LocalDate todayDate = LocalDate.now();

        // ================= SELENIUM SETUP =================
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            System.out.println("[INFO] Launching browser and opening login page...");
            driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");

            // ================= LOGIN =================
            driver.findElement(By.id("txt_username")).sendKeys("Shadowfax_cen#SUB4");
            driver.findElement(By.id("txt_password")).sendKeys("Ewaybill#1122");

            System.out.println("[ACTION REQUIRED] Enter CAPTCHA manually");
            driver.findElement(By.name("txtCaptcha")).click();

            Thread.sleep(10000); // Wait for manual Captcha entry

            driver.findElement(By.name("btnLogin")).click();
            System.out.println("[INFO] Login button clicked.");

            // ALERT HANDLE
            try {
                Thread.sleep(1500);
                Alert alert = driver.switchTo().alert();
                System.out.println("[ALERT] Login message: " + alert.getText());
                alert.accept();
            } catch (Exception e) {
                System.out.println("[INFO] No Alert presented after login");
            }

            WebElement checkbox = driver.findElement(By.id("rbOtpshow"));
            if (!checkbox.isSelected()) {
                checkbox.click();
            }
            System.out.println("[INFO] OTP Visible checkbox ticked");

            driver.findElement(By.name("OtpTxt")).click();
            System.out.println("[ACTION REQUIRED] Enter OTP manually");
            Thread.sleep(10000); // Wait for manual OTP entry

            driver.findElement(By.name("btnsubmit")).click();
            System.out.println("[INFO] Login Completed");

            // ================= PROCESSING LOOP =================

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                XSSFRow row = sheet.getRow(i);
                if (row == null) continue;

                // Column 0: E-Way Bill Number | Column 1: Remark Output
                String ewayBill = formatter.formatCellValue(row.getCell(0));

                XSSFCell remarkCell = row.getCell(1);
                if (remarkCell == null) {
                    remarkCell = row.createCell(1);
                }

                if (ewayBill.trim().isEmpty()) continue;

                System.out.println("--------------------------------------------------");
                System.out.println("[ROW " + i + "] Processing E-Way Bill: " + ewayBill);

                // Scraped attributes initialized
                String extractedCity = "";
                String extractedPincode = "";
                String extractedVehicleNo = "";

                try {
                    // ================= STEP 1: OPEN PRINT PAGE & SCRAPE DETAILS =================
                    driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");
                    Thread.sleep(1000);

                    clickMenuOption(driver, wait, js, "Print EWB");
                    Thread.sleep(500);

                    // Enter E-Way Bill Number on Print Page
                    WebElement printInput = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_txt_ebillno"))
                    );
                    printInput.clear();
                    printInput.sendKeys(ewayBill);

                    // Click GO
                    WebElement goBtn = wait.until(
                            ExpectedConditions.elementToBeClickable(By.id("ctl00_ContentPlaceHolder1_btn_go"))
                    );
                    js.executeScript("arguments[0].click();", goBtn);
                    Thread.sleep(1000);

                    // 1. Extract Pincode from Place of Dispatch span
                    try {
                        WebElement dispatchSpan = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_lbl_placeDispatc"))
                        );
                        String dispatchText = dispatchSpan.getText().trim();

                        Matcher pinMatcher = Pattern.compile("\\d{6}").matcher(dispatchText);
                        if (pinMatcher.find()) {
                            extractedPincode = pinMatcher.group();
                        }
                    } catch (Exception e) {
                        System.err.println("[WARN] Could not extract Pincode from dispatch section.");
                    }

                    // 2. Extract Vehicle Number AND Last Movement City directly from Vehicle Details Table
                    try {
                        // Vehicle Number is in Row 2, Column 2 (td[2])
                        WebElement vehicleTd = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.xpath("//table[@id='ctl00_ContentPlaceHolder1_GVVehicleDetails']//tr[2]/td[2]")
                                )
                        );
                        extractedVehicleNo = vehicleTd.getText().trim();

                        // Last Updated "From" City is in Row 2, Column 3 (td[3]) as highlighted in inspector
                        WebElement cityTd = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.xpath("//table[@id='ctl00_ContentPlaceHolder1_GVVehicleDetails']//tr[2]/td[3]")
                                )
                        );
                        extractedCity = cityTd.getText().trim();

                    } catch (Exception e) {
                        System.err.println("[WARN] Could not extract City/Vehicle Number from vehicle details table.");
                    }

                    System.out.println("[ROW " + i + " SCRAPED DATA] City: " + extractedCity + 
                                       " | Pincode: " + extractedPincode + 
                                       " | VehicleNo: " + extractedVehicleNo);

                    // ================= STEP 2: NAVIGATE TO EXTEND VALIDITY PAGE =================
                    driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");
                    Thread.sleep(1000);

                    clickMenuOption(driver, wait, js, "Extend Validity");

                    WebElement ewayBox = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_txt_no"))
                    );
                    ewayBox.clear();
                    ewayBox.sendKeys(ewayBill);

                    WebElement extendGoButton = wait.until(
                            ExpectedConditions.elementToBeClickable(By.id("ctl00_ContentPlaceHolder1_Btn_go"))
                    );
                    extendGoButton.click();
                    System.out.println("[ROW " + i + "] Clicked 'GO' on Extend Validity page.");
                    Thread.sleep(800);

                    // Check for Alert on GO (e.g. "Already Extended")
                    try {
                        Alert panelAlert = driver.switchTo().alert();
                        String alertText = panelAlert.getText();
                        panelAlert.accept();

                        System.out.println("[ROW " + i + " ALERT] " + alertText);

                        if (alertText.toLowerCase().contains("already")
                                || alertText.toLowerCase().contains("extended")
                                || alertText.toLowerCase().contains("8 hour")) {

                            remarkCell.setCellValue("Already Extended");
                            continue;
                        } else {
                            remarkCell.setCellValue("Failed: " + alertText);
                            continue;
                        }
                    } catch (org.openqa.selenium.NoAlertPresentException e) {
                        // Form loaded successfully
                    }

                    // ================= STEP 3: FILL EXTENSION FORM =================
                    // Radio button "Yes"
                    WebElement yesRadio = wait.until(
                            ExpectedConditions.presenceOfElementLocated(By.id("rbn_extent_0"))
                    );
                    js.executeScript("arguments[0].scrollIntoView(true);", yesRadio);
                    Thread.sleep(100);
                    js.executeScript("arguments[0].click();", yesRadio);

                    // Reason Dropdown -> Transhipment
                    WebElement dropdown = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$ddl_extend"))
                    );
                    Select select = new Select(dropdown);
                    select.selectByVisibleText("Transhipment");

                    // Remarks Box -> Others
                    WebElement remark = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Remarks']"))
                    );
                    remark.clear();
                    remark.sendKeys("Others");

                    // Current Place (Scraped "From" City from table)
                    WebElement txtPlace = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txt_vehFromPlace"))
                    );
                    txtPlace.clear();
                    txtPlace.sendKeys(extractedCity);

                    // Current Pincode (Scraped Pincode)
                    WebElement txtPin = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txtFromEnteredPinCode"))
                    );
                    txtPin.clear();
                    txtPin.sendKeys(extractedPincode);
                    Thread.sleep(2000);

                    // Dismiss State/Pin Alert if pop-up appears
                    try {
                        Alert pinAlert = driver.switchTo().alert();
                        pinAlert.accept();
                    } catch (org.openqa.selenium.NoAlertPresentException e) {
                    }

                    // Vehicle Number (Scraped Vehicle Number)
                    WebElement txtVehicle = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txtVehicleNo"))
                    );
                    txtVehicle.clear();
                    txtVehicle.sendKeys(extractedVehicleNo);

                    // Submit Extension Form
                    driver.findElement(By.id("btnsbmt")).click();
                    System.out.println("[ROW " + i + "] Submitted extension form.");
                    Thread.sleep(500);

                    // ================= STEP 4: DISTANCE ALERT / MODAL HANDLING =================
                    try {
                        Alert distanceAlert = driver.switchTo().alert();
                        distanceAlert.accept();
                        Thread.sleep(500);

                        WebElement txtDistance = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.xpath("//input[contains(@id,'txt_dist') or contains(@id,'txtDistance') or contains(@name,'txt_dist')]")
                                )
                        );
                        txtDistance.click();
                        txtDistance.sendKeys(Keys.CONTROL + "a");
                        txtDistance.sendKeys(Keys.BACK_SPACE);
                        txtDistance.sendKeys("10");

                        driver.findElement(By.id("btnsbmt")).click();
                        Thread.sleep(500);

                    } catch (org.openqa.selenium.NoAlertPresentException e) {
                        try {
                            WebElement okModalBtn = driver.findElement(
                                    By.xpath("//div[contains(@class,'ui-dialog-buttonset')]//button | //button[translate(text(), 'ok', 'OK')='OK']")
                            );

                            if (okModalBtn.isDisplayed()) {
                                okModalBtn.click();
                                Thread.sleep(500);

                                WebElement txtDistance = wait.until(
                                        ExpectedConditions.visibilityOfElementLocated(
                                                By.xpath("//input[contains(@id,'txt_dist') or contains(@id,'txtDistance') or contains(@name,'txt_dist')]")
                                        )
                                );
                                txtDistance.click();
                                txtDistance.sendKeys(Keys.CONTROL + "a");
                                txtDistance.sendKeys(Keys.BACK_SPACE);
                                txtDistance.sendKeys("09");

                                driver.findElement(By.id("btnsbmt")).click();
                                Thread.sleep(500);
                            }
                        } catch (Exception modalEx) {
                        }
                    }

                    // Final alert check after form submission
                    try {
                        Alert finalAlert = driver.switchTo().alert();
                        finalAlert.accept();
                    } catch (Exception alertEx) {
                    }

                    // ================= STEP 5: VERIFY UPDATED DATE VIA PRINT PAGE =================
                    driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");
                    Thread.sleep(1000);

                    clickMenuOption(driver, wait, js, "Print EWB");
                    Thread.sleep(500);

                    WebElement verifyInput = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_txt_ebillno"))
                    );
                    verifyInput.clear();
                    verifyInput.sendKeys(ewayBill);

                    WebElement verifyGoBtn = wait.until(
                            ExpectedConditions.elementToBeClickable(By.id("ctl00_ContentPlaceHolder1_btn_go"))
                    );
                    js.executeScript("arguments[0].click();", verifyGoBtn);
                    Thread.sleep(500);

                    WebElement validUntilSpan = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_lblValidTo"))
                    );

                    String validUntilText = validUntilSpan.getText().trim();
                    System.out.println("[INFO] Extracted 'Valid Until' Date: " + validUntilText);

                    Matcher matcher = Pattern.compile("\\d{2}/\\d{2}/\\d{4}").matcher(validUntilText);

                    if (matcher.find()) {
                        String extractedDateStr = matcher.group();
                        LocalDate extractedDate = LocalDate.parse(extractedDateStr, dtf);

                        if (extractedDate.equals(todayDate)) {
                            String msg = "Need To Extend (Valid Today: " + extractedDateStr + ")";
                            remarkCell.setCellValue(msg);
                            System.out.println("[ROW " + i + " STATUS] " + msg);
                        } else if (extractedDate.isAfter(todayDate)) {
                            String msg = "Extended (Valid Until: " + extractedDateStr + ")";
                            remarkCell.setCellValue(msg);
                            System.out.println("[ROW " + i + " STATUS] SUCCESS: " + msg);
                        } else {
                            String msg = "Expired (Valid Until: " + extractedDateStr + ")";
                            remarkCell.setCellValue(msg);
                            System.out.println("[ROW " + i + " STATUS] WARNING: " + msg);
                        }
                    } else {
                        remarkCell.setCellValue("Unverified Date Format: " + validUntilText);
                    }

                } catch (Exception rowException) {
                    System.err.println("[ROW " + i + " ERROR] Submission failed: " + rowException.getMessage());
                    remarkCell.setCellValue("Failed: " + rowException.getMessage());
                }
            }

        } finally {

            // ================= EXCEL WRITING =================
            try (FileOutputStream fos = new FileOutputStream(path)) {
                workbook.write(fos);
                System.out.println("[INFO] Excel sheet updated successfully.");
            } catch (Exception e) {
                System.err.println("[WARN] Excel write failed: " + e.getMessage());
            }

            if (workbook != null) workbook.close();
            driver.quit();
            System.out.println("\n[INFO] ALL E-WAY BILLS PROCESSED. BROWSER CLOSED.");
        }
    }

    // Reliable ASP.NET Menu Navigation Helper
    private static void clickMenuOption(
            WebDriver driver,
            WebDriverWait wait,
            JavascriptExecutor js,
            String menuText
    ) throws InterruptedException {

        try {
            WebElement toggleBtn = driver.findElement(
                    By.xpath("//a[contains(@class,'sidebar-toggle')] | //a[contains(@class,'nav-link')]")
            );

            if (toggleBtn.isDisplayed()) {
                js.executeScript("arguments[0].click();", toggleBtn);
                Thread.sleep(500);
            }
        } catch (Exception ignored) {
        }

        WebElement subMenuItem = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//a[contains(text(),'" + menuText + "')]")
                )
        );

        js.executeScript("arguments[0].click();", subMenuItem);
        Thread.sleep(1000);
    }
}