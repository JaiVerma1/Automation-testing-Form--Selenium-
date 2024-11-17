package com.wperest.blooddonor.tests;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.JavascriptExecutor;
import org.testng.annotations.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class FormAutomationTest {
    WebDriver driver;
    WebDriverWait wait;
    JavascriptExecutor js;
    private static final String EXCEL_PATH = "C:\\Users\\Anirudh Rautela\\eclipse-workspace\\blooddonor\\input.xlsx";
    private static final String SHEET_NAME = "Sheet1";
    private Random random = new Random();

    // Add method for random delays
    private void addHumanDelay() throws InterruptedException {
        // Random delay between 0.5 and 2 seconds
        Thread.sleep(random.nextInt(1500) + 500);
    }

    // Add method for human-like typing
    private void typeHumanlike(WebElement element, String text) throws InterruptedException {
        for (char c : text.toCharArray()) {
            element.sendKeys(String.valueOf(c));
            // Random typing delay between 50 and 150 milliseconds
            Thread.sleep(random.nextInt(100) + 50);
        }
    }

    @BeforeMethod
    public void setup() {
        driver = new EdgeDriver();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        js = (JavascriptExecutor) driver;
        
        if (driver == null) {
            System.out.println("WebDriver is null after initialization!");
        } else {
            System.out.println("WebDriver initialized successfully.");
        }
        driver.manage().window().maximize();
        driver.get("https://qavalidation.com/demo-form/");
    }

    @Test(dataProvider = "excelFormData")
    public void testFormSubmission(String name, String email, String phone, String gender, String experience,
                                   String skills, String qaTool, String otherDetails) throws InterruptedException {
        // Wait for form to be fully loaded
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//form")));
        addHumanDelay();

        // Full Name
        WebElement nameField = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[1]/input")));
        typeHumanlike(nameField, name);
        addHumanDelay();
        
        // Email
        WebElement emailField = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[2]/input")));
        typeHumanlike(emailField, email);
        addHumanDelay();
        
        // Phone Number
        WebElement phoneField = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[3]/input")));
        typeHumanlike(phoneField, phone);
        addHumanDelay();
        
        // Gender Dropdown
        WebElement genderDropdown = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[4]/div/select")));
        try {
            new Select(genderDropdown).selectByVisibleText(gender);
            addHumanDelay();
        } catch (Exception e) {
            new Select(genderDropdown).selectByIndex(0);
            addHumanDelay();
        }

        // Years of Experience Radio Button
        String experienceXPath = "";
        switch (experience) {
            case "1": experienceXPath = "p[1]"; break;
            case "2": experienceXPath = "p[2]"; break;
            case "3": experienceXPath = "p[3]"; break;
            case "4": experienceXPath = "p[4]"; break;
            case "5": experienceXPath = "p[5]"; break;
            case "Above 5": experienceXPath = "p[6]"; break;
            default: throw new IllegalArgumentException("Invalid experience: " + experience);
        }
        
        WebElement expRadio = wait.until(ExpectedConditions.presenceOfElementLocated(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[5]/fieldset/" + experienceXPath + "/input")));
        js.executeScript("arguments[0].scrollIntoView(true);", expRadio);
        addHumanDelay();
        js.executeScript("arguments[0].click();", expRadio);
        addHumanDelay();

        // Skills Checkboxes
        String[] skillsArray = skills.split(",");
        for (String skill : skillsArray) {
            String skillXPath = "";
            switch (skill.trim()) {
                case "Functional Testing": skillXPath = "p[1]"; break;
                case "Automation Testing": skillXPath = "p[2]"; break;
                case "API Testing": skillXPath = "p[3]"; break;
                case "DB Testing": skillXPath = "p[4]"; break;
                default: throw new IllegalArgumentException("Invalid skill: " + skill);
            }
            
            WebElement skillCheckbox = wait.until(ExpectedConditions.presenceOfElementLocated(
                By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[6]/fieldset/" + skillXPath + "/input")));
            js.executeScript("arguments[0].scrollIntoView(true);", skillCheckbox);
            addHumanDelay();
            js.executeScript("arguments[0].click();", skillCheckbox);
            addHumanDelay();
        }

        // QA Tools Dropdown
        WebElement qaToolDropdown = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[7]/div/select")));
        new Select(qaToolDropdown).selectByVisibleText(qaTool);
        addHumanDelay();

        // Other Details
        WebElement detailsField = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[8]/textarea")));
        typeHumanlike(detailsField, otherDetails);
        addHumanDelay();

        // Submit Form
        WebElement submitButton = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("/html/body/div[1]/div[5]/div[2]/div/div/div/form/div[9]/button")));
        js.executeScript("arguments[0].click();", submitButton);

        // Longer wait after submission
        Thread.sleep(random.nextInt(2000) + 2000);
    }

    @DataProvider(name = "excelFormData")
    public Object[][] getFormData() throws IOException {
        FileInputStream fis = new FileInputStream(EXCEL_PATH);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(SHEET_NAME);
        
        // Get physical number of rows (excluding empty rows)
        int lastRowNum = sheet.getLastRowNum();
        int physicalNumberOfRows = 0;
        for (int i = 1; i <= lastRowNum; i++) {
            if (sheet.getRow(i) != null && !isRowEmpty(sheet.getRow(i))) {
                physicalNumberOfRows++;
            }
        }
        
        List<Object[]> testData = new ArrayList<>();
        
        // Start from row 1 to skip header, only process up to 10 rows
        for (int i = 1; i <= Math.min(10, physicalNumberOfRows); i++) {
            Row row = sheet.getRow(i);
            if (row != null && !isRowEmpty(row)) {
                Object[] rowData = new Object[8];
                for (int j = 0; j < 8; j++) {
                    Cell cell = row.getCell(j);
                    if (j == 2) { // Phone number column
                        rowData[j] = getPhoneNumberAsString(cell);
                    } else {
                        rowData[j] = getCellValueAsString(cell);
                    }
                }
                testData.add(rowData);
            }
        }
        
        workbook.close();
        fis.close();
        
        return testData.toArray(new Object[0][]);
    }
    
    // Add method to check if row is empty
    private boolean isRowEmpty(Row row) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }
    
    // Add specific method for phone number handling
    private String getPhoneNumberAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // Format number to preserve leading zeros and full number
                long phoneNumber = (long) cell.getNumericCellValue();
                return String.format("%010d", phoneNumber); // Ensures 10 digits with leading zeros
            default:
                return "";
        }
    }
    
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    @AfterMethod
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }
    }
}