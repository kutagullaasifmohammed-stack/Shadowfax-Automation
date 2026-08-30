package Automation;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
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
        } catch (Exception e) {
            System.err.println("CRITICAL ERROR: Make sure 'AutoExtendFile.xlsx' is CLOSED in Excel before running!");
            throw e;
        } finally {
            if (fis != null) fis.close();
        }

        XSSFSheet sheet = workbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter(); 

        // ================= SELENIUM SETUP =================
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10)); // Reduced duration for faster failure detection

        driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");

        // ================= LOGIN =================
        driver.findElement(By.id("txt_username")).sendKeys("Shadowfax_cen#SUB4");
        driver.findElement(By.id("txt_password")).sendKeys("Ewaybill#2026");

        System.out.println("Enter CAPTCHA manually");
        driver.findElement(By.name("txtCaptcha")).click();

        // CAPTCHA TIME
        Thread.sleep(10000);

        // CLICK LOGIN
        driver.findElement(By.name("btnLogin")).click();

        // ALERT HANDLE
        try {
            Thread.sleep(1500);
            Alert alert = driver.switchTo().alert();
            alert.accept();
        } catch (Exception e) {
            System.out.println("No Alert presented after login");
        }

        // Tick Show OTP checkbox
        WebElement checkbox = driver.findElement(By.id("rbOtpshow"));
        if (!checkbox.isSelected()) {
            checkbox.click();
        }
        System.out.println("OTP Visible checkbox ticked");

        // OTP TIME
        driver.findElement(By.name("OtpTxt")).click();
        Thread.sleep(10000);

        // OTP SUBMIT
        driver.findElement(By.name("btnsubmit")).click();
        System.out.println("Login Completed");

        // ================= OPEN EXTEND =================
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(),'e-Waybill')]"))).click();
        System.out.println("E-way Bill Menu Clicked");

        WebElement extend = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Extend Validity')]")));
        extend.click();
        System.out.println("Extend Page Opened");

        // ================= PROCESSING LOOP =================
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            
            XSSFRow row = sheet.getRow(i);
            if (row == null) continue; 

            XSSFCell remarkCell = row.getCell(4);
            if (remarkCell == null) {
                remarkCell = row.createCell(4);
            }

            String ewayBill   = formatter.formatCellValue(row.getCell(0));
            String vehicleNo  = formatter.formatCellValue(row.getCell(1));
            String city       = formatter.formatCellValue(row.getCell(2));
            String pincode    = formatter.formatCellValue(row.getCell(3));

            if (ewayBill.trim().isEmpty()) continue;

            System.out.println("Processing Row " + i + " -> E-way Bill: " + ewayBill);

            boolean isAlreadyHandled = false;

            try {
                // 1. Enter E-way Bill
                WebElement ewayBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_ContentPlaceHolder1_txt_no")));
                ewayBox.clear();
                ewayBox.sendKeys(ewayBill);

                // 2. Click GO
                driver.findElement(By.id("ctl00_ContentPlaceHolder1_Btn_go")).click();
                Thread.sleep(100); 

                // --- FAST DETECTION: Initial Validation Error Alert (e.g., Already extended / 8 hour restriction) ---
                try {
                    Alert panelAlert = driver.switchTo().alert();
                    String alertText = panelAlert.getText();
                    panelAlert.accept(); 
                    System.out.println("Fast-skipped row " + i + " due to Portal Error: " + alertText);
                    remarkCell.setCellValue("Failed: " + alertText);
                    isAlreadyHandled = true; 
                    
                    // Clear field values via manual string wipe if needed, or proceed directly
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_txt_no")).clear();
                    continue; // Immediately jumps to next loop index, avoiding long reloads
                } catch (org.openqa.selenium.NoAlertPresentException e) {
                    // Path clear, bill valid to proceed
                }

                // 3. Click YES for Extension
                WebElement yesRadio = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("rbn_extent_0")));
                JavascriptExecutor js = (JavascriptExecutor) driver;

                js.executeScript("arguments[0].scrollIntoView(true);", yesRadio);
                Thread.sleep(100);
                js.executeScript("arguments[0].click();", yesRadio);

                // 4. Reason Dropdown Selection
                WebElement dropdown = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$ddl_extend")));
                Select select = new Select(dropdown);
                select.selectByVisibleText("Transhipment");

                // 5. Remarks Box
                WebElement remark = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Remarks']")));
                remark.clear();
                remark.sendKeys("Others");

                // 6. Current Place (City)
                WebElement txtPlace = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txt_vehFromPlace")));
                txtPlace.clear();
                txtPlace.sendKeys(city);

                // 7. Current Pincode
                WebElement txtPin = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txtFromEnteredPinCode")));
                txtPin.clear();
                txtPin.sendKeys(pincode);
                Thread.sleep(100); 

                // --- CHECK FOR PINCODE STATE ALERT ---
                try {
                    Alert pinAlert = driver.switchTo().alert();
                    pinAlert.accept(); 
                } catch (org.openqa.selenium.NoAlertPresentException e) {
                    // Path clear
                }

                // 8. Vehicle Number
                WebElement txtVehicle = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("ctl00$ContentPlaceHolder1$txtVehicleNo")));
                txtVehicle.clear();
                txtVehicle.sendKeys(vehicleNo);

                // 9. SUBMIT THE EXTENSION FORM
                driver.findElement(By.id("btnsbmt")).click();
                Thread.sleep(100); 

                // --- DETECT AND FIX BOTH DISTANCE ALERTS ---
                try {
                    Alert distanceAlert = driver.switchTo().alert();
                    String alertMsg = distanceAlert.getText();
                    
                    if (alertMsg.toLowerCase().contains("distance")) {
                        distanceAlert.accept(); 
                        Thread.sleep(100);
                        
                        WebElement txtDistance = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("txtDistance")));
                        txtDistance.clear();
                        txtDistance.sendKeys("20"); 
                        
                        driver.findElement(By.id("btnsbmt")).click();
                        Thread.sleep(100);
                    } else {
                        distanceAlert.accept();
                    }
                } catch (org.openqa.selenium.NoAlertPresentException e) {
                    // Path clear
                }

                // Final post-processing validation response check
                try {
                    Alert finalAlert = driver.switchTo().alert();
                    finalAlert.accept();
                } catch (Exception alertEx) {
                    // No final dialog message
                }

                remarkCell.setCellValue("Extended Successfully");
                System.out.println("Completed Extension Request for Bill: " + ewayBill);

            } catch (Exception rowException) {
                if (!isAlreadyHandled) {
                    System.out.println("Failed to extend row " + i + " Reason: " + rowException.getMessage());
                    remarkCell.setCellValue("Failed: " + rowException.getMessage());
                }
            }

            // 10. FAST RESET USING THE RED EXIT BUTTON
            try {
                // Clicks the Red EXIT button shown on your UI next to the Go button to clear the page fast.
                WebElement exitButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Exit')]")));
                exitButton.click();
                Thread.sleep(100);
            } catch (Exception navigationException) {
                // Fallback option if the Exit button is not present or form layout changes dynamically
                driver.get("https://ewaybillgst.gov.in/MainMenu.aspx");
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(),'e-Waybill')]"))).click();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Extend Validity')]"))).click();
            }
        }

        // ================= EXCEL WRITING =================
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        
        fos.close();
        workbook.close();

        System.out.println("ALL E-WAY BILLS PROCESSED. EXCEL SHEET UPDATED.");
        driver.quit();
    }
}
