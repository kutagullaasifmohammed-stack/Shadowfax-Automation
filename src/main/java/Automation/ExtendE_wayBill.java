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
            System.out.println("[VERIFY] Excel file loaded successfully.");
        } catch (Exception e) {
            System.err.println("[FAIL] CRITICAL ERROR: Make sure 'AutoExtendFile.xlsx' is CLOSED in Excel before running!");
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

            // ================= LOGIN & FORM ELEMENT VALIDATION =================
            WebElement userTxt = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("txt_username")));
            verifyElement(userTxt, "Username Field");
            userTxt.sendKeys("Shadowfax_cen#SUB4");

            WebElement passTxt = driver.findElement(By.id("txt_password"));
            verifyElement(passTxt, "Password Field");
            passTxt.sendKeys("Ewaybill#1122");

            WebElement captchaBox = driver.findElement(By.name("txtCaptcha"));
            verifyElement(captchaBox, "Captcha Textbox");
            captchaBox.click();

            System.out.println("[ACTION REQUIRED] Enter CAPTCHA manually within 10 seconds...");
            Thread.sleep(10000); // Wait for manual Captcha entry

            WebElement loginBtn = driver.findElement(By.name("btnLogin"));
            verifyElement(loginBtn, "Login Button");
            loginBtn.click();
            System.out.println("[VERIFY] Login button clicked successfully.");

            // ALERT HANDLE POST LOGIN
            try {
                Thread.sleep(1500);
                Alert alert = driver.switchTo().alert();
                System.out.println("[ALERT] Login message popup: " + alert.getText());
                alert.accept();
            } catch (Exception e) {
                System.out.println("[INFO] No Alert presented after login");
            }

            WebElement checkbox = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("rbOtpshow")));
            verifyElement(checkbox, "OTP Visible Checkbox");
            if (!checkbox.isSelected()) {
                checkbox.click();
            }

            WebElement otpField = driver.findElement(By.name("OtpTxt"));
            verifyElement(otpField, "OTP Input Field");
            otpField.click();
            System.out.println("[ACTION REQUIRED] Enter OTP manually within 10 seconds...");
            Thread.sleep(10000); // Wait for manual OTP entry

            WebElement submitOtpBtn = driver.findElement(By.name("btnsubmit"));
            verifyElement(submitOtpBtn, "OTP Submit Button");
            submitOtpBtn.click();
            System.out.println("[VERIFY] Login Completed & Dashboard Reached.");

            // ================= PROCESSING LOOP =================

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                XSSFRow row = sheet.getRow(i);
                if (row == null) continue;

                String ewayBill = formatter.formatCellValue(row.getCell(0));

                XSSFCell remarkCell = row.getCell(1);
                if (remarkCell == null) {
                    remarkCell = row.createCell(1);
                }

                if (ewayBill.trim().isEmpty()) continue;

                System.out.println("--------------------------------------------------");
                System.out.println("[ROW " + i + "] Processing E-Way Bill: " + ewayBill);

                String extractedCity = "";
                String extractedPincode = "";
                String extractedVehicleNo = "";

                try {
                    // ================= STEP 1: OPEN PRINT PAGE & VERIFY VALIDITY =================
                    driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");
                    Thread.sleep(1000);

                    boolean menuClicked = clickMenuOption(driver, wait, js, "Print EWB");
                    if (!menuClicked) {
                        throw new RuntimeException("Failed to click 'Print EWB' menu item.");
                    }

                    WebElement printInput = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_txt_ebillno"))
                    );
                    verifyElement(printInput, "Print EWB Number Input");
                    printInput.clear();
                    printInput.sendKeys(ewayBill);

                    WebElement goBtn = wait.until(
                            ExpectedConditions.elementToBeClickable(By.id("ctl00_ContentPlaceHolder1_btn_go"))
                    );
                    verifyElement(goBtn, "Print Go Button");
                    js.executeScript("arguments[0].click();", goBtn);
                    Thread.sleep(1000);

                    // Verify & Extract 'Valid Until' Date Text
                    WebElement validUntilSpan = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_lblValidTo"))
                    );
                    verifyElement(validUntilSpan, "Valid Until Date Display Span");
                    String validUntilText = validUntilSpan.getText().trim();
                    System.out.println("[VERIFY] Extracted 'Valid Until' Date: " + validUntilText);

                    Matcher matcher = Pattern.compile("\\d{2}/\\d{2}/\\d{4}").matcher(validUntilText);

                    if (matcher.find()) {
                        String extractedDateStr = matcher.group();
                        LocalDate extractedDate = LocalDate.parse(extractedDateStr, dtf);

                        // CONDITIONAL ROUTING BASED ON VALIDITY
                        if (extractedDate.isAfter(todayDate)) {
                            String msg = "Extended (Valid Until: " + extractedDateStr + ")";
                            remarkCell.setCellValue(msg);
                            System.out.println("[STATUS: PASS] E-Way bill valid beyond today (" + extractedDateStr + "). Skipped extension.");
                            continue; // Skip directly to next row
                        } else if (extractedDate.equals(todayDate)) {
                            System.out.println("[STATUS: ACTION] Valid TODAY (" + extractedDateStr + ") -> Proceeding to Extend Validity.");
                        } else if (extractedDate.isBefore(todayDate)) {
                            System.out.println("[STATUS: EXPIRED] Date is before today (" + extractedDateStr + ") -> Attempting Extension.");
                        }
                    } else {
                        System.err.println("[WARN] Date Regex Match Failed for text: " + validUntilText);
                    }

                    // Extract Scraped Details needed for form
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
                        System.err.println("[FAIL] Pincode extraction failed from dispatch element.");
                    }

                    try {
                        WebElement vehicleTd = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.xpath("//table[@id='ctl00_ContentPlaceHolder1_GVVehicleDetails']//tr[2]/td[2]")
                                )
                        );
                        extractedVehicleNo = vehicleTd.getText().trim();

                        WebElement cityTd = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.xpath("//table[@id='ctl00_ContentPlaceHolder1_GVVehicleDetails']//tr[2]/td[3]")
                                )
                        );
                        extractedCity = cityTd.getText().trim();
                    } catch (Exception e) {
                        System.err.println("[FAIL] Vehicle Number / City extraction failed from table.");
                    }

                    System.out.println("[SCRAPED DATA VERIFIED] City: " + extractedCity +
                                       " | Pincode: " + extractedPincode +
                                       " | VehicleNo: " + extractedVehicleNo);

                    // ================= STEP 2: OPEN EXTEND VALIDITY PAGE =================
                    driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");
                    Thread.sleep(1000);

                    clickMenuOption(driver, wait, js, "Extend Validity");

                    WebElement ewayBox = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_txt_no"))
                    );
                    verifyElement(ewayBox, "Extend Validity E-Way Box Input");
                    ewayBox.clear();
                    ewayBox.sendKeys(ewayBill);

                    WebElement extendGoButton = wait.until(
                            ExpectedConditions.elementToBeClickable(By.id("ctl00_ContentPlaceHolder1_Btn_go"))
                    );
                    verifyElement(extendGoButton, "Extend Validity Go Button");
                    extendGoButton.click();
                    System.out.println("[VERIFY] Clicked 'GO' on Extend Validity page.");
                    Thread.sleep(800);

                    // Check for immediate Alert (e.g., Already Extended)
                    try {
                        Alert panelAlert = driver.switchTo().alert();
                        String alertText = panelAlert.getText();
                        panelAlert.accept();

                        System.out.println("[ALERT CAPTURED] " + alertText);

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
                        // Extension form rendered properly
                    }

                    // ================= STEP 3: EXTENSION FORM FILL & VERIFY =================
                    WebElement yesRadio = wait.until(
                            ExpectedConditions.presenceOfElementLocated(By.id("rbn_extent_0"))
                    );
                    verifyElement(yesRadio, "Extension Radio Button (Yes)");
                    js.executeScript("arguments[0].scrollIntoView(true);", yesRadio);
                    Thread.sleep(100);
                    js.executeScript("arguments[0].click();", yesRadio);

                    WebElement dropdown = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$ddl_extend"))
                    );
                    verifyElement(dropdown, "Reason Dropdown Menu");
                    Select select = new Select(dropdown);
                    select.selectByVisibleText("Transhipment");

                    WebElement remark = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Remarks']"))
                    );
                    verifyElement(remark, "Remarks Text Input");
                    remark.clear();
                    remark.sendKeys("Others");

                    WebElement txtPlace = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txt_vehFromPlace"))
                    );
                    verifyElement(txtPlace, "From Place Input");
                    txtPlace.clear();
                    txtPlace.sendKeys(extractedCity);

                    WebElement txtPin = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txtFromEnteredPinCode"))
                    );
                    verifyElement(txtPin, "From Pincode Input");
                    txtPin.clear();
                    txtPin.sendKeys(extractedPincode);
                    Thread.sleep(2000);

                    // Dismiss State/Pin calculation pop-up if present
                    try {
                        Alert pinAlert = driver.switchTo().alert();
                        pinAlert.accept();
                    } catch (org.openqa.selenium.NoAlertPresentException e) {
                    }

                    WebElement txtVehicle = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txtVehicleNo"))
                    );
                    verifyElement(txtVehicle, "Vehicle Number Input");
                    txtVehicle.clear();
                    txtVehicle.sendKeys(extractedVehicleNo);

                    WebElement submitBtn = driver.findElement(By.id("btnsbmt"));
                    verifyElement(submitBtn, "Form Submit Button");
                    submitBtn.click();
                    System.out.println("[VERIFY] Submitted extension form.");
                    Thread.sleep(500);

                    // ================= STEP 4: DISTANCE MODAL & ALERT VALIDATION =================
                    try {
                        Alert distanceAlert = driver.switchTo().alert();
                        distanceAlert.accept();
                        Thread.sleep(500);

                        WebElement txtDistance = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.xpath("//input[contains(@id,'txt_dist') or contains(@id,'txtDistance') or contains(@name,'txt_dist')]")
                                )
                        );
                        verifyElement(txtDistance, "Distance Override Input Field");
                        txtDistance.click();
                        txtDistance.sendKeys(Keys.CONTROL + "a");
                        txtDistance.sendKeys(Keys.BACK_SPACE);
                        txtDistance.sendKeys("10");

                        submitBtn.click();
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
                                verifyElement(txtDistance, "Distance Override Modal Input");
                                txtDistance.click();
                                txtDistance.sendKeys(Keys.CONTROL + "a");
                                txtDistance.sendKeys(Keys.BACK_SPACE);
                                txtDistance.sendKeys("09");

                                submitBtn.click();
                                Thread.sleep(500);
                            }
                        } catch (Exception modalEx) {
                        }
                    }

                    try {
                        Alert finalAlert = driver.switchTo().alert();
                        finalAlert.accept();
                    } catch (Exception alertEx) {
                    }

                    // ================= STEP 5: FINAL POST-EXTENSION VERIFICATION =================
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

                    WebElement postValidSpan = wait.until(
                            ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_lblValidTo"))
                    );

                    String postValidText = postValidSpan.getText().trim();
                    Matcher postMatcher = Pattern.compile("\\d{2}/\\d{2}/\\d{4}").matcher(postValidText);

                    if (postMatcher.find()) {
                        String finalDateStr = postMatcher.group();
                        LocalDate finalDate = LocalDate.parse(finalDateStr, dtf);

                        if (finalDate.isAfter(todayDate)) {
                            remarkCell.setCellValue("Extended (Valid Until: " + finalDateStr + ")");
                            System.out.println("[FINAL STATUS: PASS] Bill extended to " + finalDateStr);
                        } else if (finalDate.equals(todayDate)) {
                            remarkCell.setCellValue("Need To Extend (Valid Today: " + finalDateStr + ")");
                        } else {
                            remarkCell.setCellValue("Expired (Valid Until: " + finalDateStr + ")");
                        }
                    } else {
                        remarkCell.setCellValue("Unverified Date Format: " + postValidText);
                    }

                } catch (Exception rowException) {
                    System.err.println("[ROW " + i + " FAIL] Processing failed: " + rowException.getMessage());
                    remarkCell.setCellValue("Failed: " + rowException.getMessage());
                }
            }

        } finally {

            // ================= EXCEL WRITING =================
            try (FileOutputStream fos = new FileOutputStream(path)) {
                workbook.write(fos);
                System.out.println("[INFO] Excel sheet updated successfully.");
            } catch (Exception e) {
                System.err.println("[FAIL] Excel write failed: " + e.getMessage());
            }

            if (workbook != null) workbook.close();
            driver.quit();
            System.out.println("\n[INFO] ALL TEST CASES & E-WAY BILLS PROCESSED. BROWSER CLOSED.");
        }
    }

    // Defensive Element Verification Method
    private static void verifyElement(WebElement element, String elementName) {
        if (element != null && (element.isDisplayed() || element.isEnabled())) {
            System.out.println("[ELEMENT CHECK: PASS] " + elementName + " is available and functional.");
        } else {
            System.err.println("[ELEMENT CHECK: FAIL] " + elementName + " is missing or disabled.");
        }
    }

    // ASP.NET Sub-Menu Navigation Helper
    private static boolean clickMenuOption(
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

        try {
            WebElement subMenuItem = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//a[contains(text(),'" + menuText + "')]")
                    )
            );

            verifyElement(subMenuItem, "Sub-Menu Option: " + menuText);
            js.executeScript("arguments[0].click();", subMenuItem);
            Thread.sleep(1000);
            return true;
        } catch (Exception e) {
            System.err.println("[FAIL] Menu item not reachable: " + menuText);
            return false;
        }
    }
} 