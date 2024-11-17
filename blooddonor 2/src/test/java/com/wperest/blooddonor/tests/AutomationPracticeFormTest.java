package com.wperest.blooddonor.tests;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Random;

public class AutomationPracticeFormTest {

    WebDriver driver;
    WebDriverWait wait;
    Random random = new Random();
    
    @BeforeClass
    public void setup() {
        driver = new EdgeDriver();
        // Reduced wait time from 10 to 5 seconds
        wait = new WebDriverWait(driver, Duration.ofSeconds(5));
        driver.manage().window().maximize();
    }

    private void addRandomDelay() {
        try {
            // Reduced delay to between 500ms to 1.5 seconds
            Thread.sleep(random.nextInt(1000) + 500);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }

    @Test(dataProvider = "excelData")
    public void fillForm(String firstName, String lastName, String email, String gender, 
                        String mobile, String dob, String subjects, String hobbies, 
                        String picturePath, String address, String state, String city) {
        try {
            driver.get("https://demoqa.com/automation-practice-form");
            wait.until(ExpectedConditions.elementToBeClickable(By.id("firstName")));
            addRandomDelay();

            // Fill basic information
            fillField(By.id("firstName"), firstName);
            fillField(By.id("lastName"), lastName);
            fillField(By.id("userEmail"), email);
            fillField(By.id("userNumber"), mobile);

            // Handle gender
            selectGender(gender);

            // Handle date of birth
            handleDateOfBirth(dob);

            // Handle subjects
            handleSubjects(subjects);

            // Handle hobbies
            selectHobbies(hobbies);

            // Handle file upload
            if (picturePath != null && !picturePath.trim().isEmpty()) {
                WebElement uploadElement = driver.findElement(By.id("uploadPicture"));
                uploadElement.sendKeys(picturePath);
            }

            // Fill address
            fillField(By.id("currentAddress"), address);

            // Handle state and city
            handleStateAndCity(state, city);

            // Submit form
            JavascriptExecutor js = (JavascriptExecutor) driver;
            WebElement submitButton = driver.findElement(By.id("submit"));
            js.executeScript("arguments[0].scrollIntoView(true);", submitButton);
            js.executeScript("arguments[0].click();", submitButton);

            // Handle submission modal with specific XPath
            handleSubmissionModal();

        } catch (Exception e) {
            System.err.println("Error filling form: " + e.getMessage());
            throw e;
        }
    }

    private void fillField(By locator, String value) {
        if (value != null && !value.trim().isEmpty()) {
            WebElement element = wait.until(ExpectedConditions.elementToBeClickable(locator));
            element.clear();
            // Reduced typing delay to 25-75ms
            for (char c : value.toCharArray()) {
                element.sendKeys(String.valueOf(c));
                try {
                    Thread.sleep(random.nextInt(50) + 25);
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                }
            }
        }
    }

    private void selectGender(String gender) {
        if (gender != null && !gender.trim().isEmpty()) {
            String genderXPath = String.format("//label[text()='%s']", gender);
            WebElement genderElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(genderXPath)));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", genderElement);
        }
    }

    private String standardizeDateFormat(Cell cell) {
        try {
            if (cell.getCellType() == CellType.STRING) {
                // If it's already in string format like "21-02-1996"
                return cell.getStringCellValue().trim();
            } else if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                // If it's in Excel date format
                java.util.Date date = cell.getDateCellValue();
                return new java.text.SimpleDateFormat("dd-MM-yyyy").format(date);
            }
        } catch (Exception e) {
            System.err.println("Error processing date: " + e.getMessage());
        }
        return "";
    }

    // Then, update the calendar interaction method
    private void handleDateOfBirth(String dob) {
        if (dob != null && !dob.trim().isEmpty()) {
            try {
                // Parse the date
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
                LocalDate date = LocalDate.parse(dob, formatter);
                
                // Click the date input field
                WebElement dateField = wait.until(ExpectedConditions.elementToBeClickable(
                    By.id("dateOfBirthInput")));
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", dateField);
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", dateField);
                addRandomDelay();

                // Select year first
                WebElement yearDropdown = wait.until(ExpectedConditions.elementToBeClickable(
                    By.className("react-datepicker__year-select")));
                Select yearSelect = new Select(yearDropdown);
                yearSelect.selectByValue(String.valueOf(date.getYear()));
                addRandomDelay();

                // Select month
                WebElement monthDropdown = wait.until(ExpectedConditions.elementToBeClickable(
                    By.className("react-datepicker__month-select")));
                Select monthSelect = new Select(monthDropdown);
                monthSelect.selectByValue(String.valueOf(date.getMonthValue() - 1)); // Months are 0-based
                addRandomDelay();

                // Select day - using a more robust XPath
                String dayXPath = String.format(
                    "//div[contains(@class, 'react-datepicker__day') and " +
                    "not(contains(@class, 'react-datepicker__day--outside-month')) and " +
                    "contains(@class, 'react-datepicker__day--0%d') and " +
                    "not(contains(@class, 'react-datepicker__day--selected'))]", 
                    date.getDayOfMonth());
                
                WebElement dayElement = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath(dayXPath)));
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", dayElement);
                
                // Add debug logging
                System.out.println("Date selected: " + dob);
                
            } catch (Exception e) {
                System.err.println("Error selecting date: " + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    private void handleSubjects(String subjects) {
        if (subjects != null && !subjects.trim().isEmpty()) {
            String[] subjectList = subjects.split(",");
            for (String subject : subjectList) {
                WebElement subjectInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.id("subjectsInput")));
                subject = subject.trim();
                // Type subject character by character
                for (char c : subject.toCharArray()) {
                    subjectInput.sendKeys(String.valueOf(c));
                    try {
                        Thread.sleep(random.nextInt(150) + 50);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                    }
                }
                addRandomDelay();
                
                // Wait for and click the suggestion
                WebElement suggestion = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//div[contains(@class, 'subjects-auto-complete__option')]")));
                suggestion.click();
                addRandomDelay();
            }
        }
    }

    private void selectHobbies(String hobbies) {
        if (hobbies != null && !hobbies.trim().isEmpty()) {
            String[] hobbyList = hobbies.split(",");
            for (String hobby : hobbyList) {
                hobby = hobby.trim();
                // Map the hobby text to the exact label text on the form
                String mappedHobby;
                switch (hobby.toLowerCase()) {
                    case "sports":
                        mappedHobby = "Sports";
                        break;
                    case "reading":
                        mappedHobby = "Reading";
                        break;
                    case "music":
                        mappedHobby = "Music";
                        break;
                    default:
                        System.out.println("Unrecognized hobby: " + hobby);
                        continue;
                }
                
                try {
                    String hobbyXPath = String.format("//label[text()='%s']", mappedHobby);
                    WebElement hobbyElement = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath(hobbyXPath)));
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", 
                        hobbyElement);
                    addRandomDelay();
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", 
                        hobbyElement);
                    addRandomDelay();
                } catch (Exception e) {
                    System.err.println("Unable to select hobby: " + hobby + ". Error: " + 
                        e.getMessage());
                }
            }
        }
    }

    private void handleStateAndCity(String state, String city) {
        if (state != null && !state.trim().isEmpty()) {
            try {
                // Improved state selection with retries
                int maxRetries = 3;
                int retryCount = 0;
                boolean stateSelected = false;
                
                while (!stateSelected && retryCount < maxRetries) {
                    try {
                        WebElement stateDropdown = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("#state .css-1hwfws3")));
                        ((JavascriptExecutor) driver).executeScript(
                            "arguments[0].scrollIntoView(true);", stateDropdown);
                        addRandomDelay();
                        stateDropdown.click();
                        addRandomDelay();

                        WebElement stateOption = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath(String.format(
                                "//div[contains(@class, ' css-26l3qy-menu')]//div[text()='%s']", 
                                state))));
                        stateOption.click();
                        stateSelected = true;
                    } catch (Exception e) {
                        retryCount++;
                        if (retryCount == maxRetries) {
                            throw e;
                        }
                        addRandomDelay();
                    }
                }

                // Handle city selection only if state was successfully selected
                if (city != null && !city.trim().isEmpty() && stateSelected) {
                    addRandomDelay(); // Wait for city dropdown to become active
                    
                    WebElement cityDropdown = wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("#city .css-1hwfws3")));
                    ((JavascriptExecutor) driver).executeScript(
                        "arguments[0].scrollIntoView(true);", cityDropdown);
                    addRandomDelay();
                    cityDropdown.click();
                    addRandomDelay();

                    WebElement cityOption = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath(String.format(
                            "//div[contains(@class, ' css-26l3qy-menu')]//div[text()='%s']", 
                            city))));
                    cityOption.click();
                    addRandomDelay();
                }
            } catch (Exception e) {
                System.err.println("Error handling state/city selection: " + e.getMessage());
                throw e;
            }
        }
    }

     private void handleSubmissionModal() {
        try {
            // Wait for modal to appear
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("modal-content")));
            
            // Use the specific XPath for the close button
            WebElement closeButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("/html/body/div[4]/div/div/div[3]/button")));
            
            // Ensure modal is fully loaded before clicking
            addRandomDelay();
            
            // Try clicking with both regular click and JavaScript executor
            try {
                closeButton.click();
            } catch (ElementClickInterceptedException e) {
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].click();", closeButton);
            }
            
            // Wait for modal to completely disappear
            wait.until(ExpectedConditions.invisibilityOfElementLocated(By.className("modal-content")));
            
            // Additional delay to ensure the page is ready for the next entry
            addRandomDelay();
            
        } catch (TimeoutException e) {
            System.err.println("Modal dialog not found or already closed: " + e.getMessage());
            // Take screenshot or add additional error handling if needed
            throw e;
        }
    }

     @DataProvider(name = "excelData")
     public Object[][] readExcel() throws IOException {
         String excelFilePath = "C:\\Users\\Anirudh Rautela\\eclipse-workspace\\blooddonor\\input2.xlsx";
         FileInputStream fileInputStream = new FileInputStream(excelFilePath);
         Workbook workbook = new XSSFWorkbook(fileInputStream);
         Sheet sheet = workbook.getSheetAt(0);
         
         int rowCount = Math.min(sheet.getPhysicalNumberOfRows(), 10);
         Object[][] data = new Object[9][12];
         
         for (int i = 1; i < rowCount; i++) {
             Row row = sheet.getRow(i);
             for (int j = 0; j < 12; j++) {
                 Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                 
                 // Special handling for the date column (assuming it's column index 5)
                 if (j == 5) {
                     data[i - 1][j] = standardizeDateFormat(cell);
                     // Debug log
                     System.out.println("Read date from Excel: " + data[i - 1][j]);
                 } else {
                     // Handle other columns as before
                     switch (cell.getCellType()) {
                         case NUMERIC:
                             double numericValue = cell.getNumericCellValue();
                             if (numericValue == (long) numericValue) {
                                 data[i - 1][j] = String.format("%.0f", numericValue);
                             } else {
                                 data[i - 1][j] = String.valueOf(numericValue);
                             }
                             break;
                         case STRING:
                             data[i - 1][j] = cell.getStringCellValue().trim();
                             break;
                         case FORMULA:
                             try {
                                 data[i - 1][j] = String.valueOf(cell.getNumericCellValue());
                             } catch (IllegalStateException e) {
                                 data[i - 1][j] = cell.getStringCellValue().trim();
                             }
                             break;
                         default:
                             data[i - 1][j] = "";
                     }
                 }
             }
         }
         workbook.close();
         fileInputStream.close();
         return data;
     }
}