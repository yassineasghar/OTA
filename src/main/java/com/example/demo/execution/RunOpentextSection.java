package com.example.demo.execution;

import com.example.demo.OpenTextAutomation;
import com.example.demo.exception.TestFailedException;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static com.example.demo.execution.TimeConstants.*;

@Component
public class RunOpentextSection {

    private final Logger LOG = LoggerFactory.getLogger(RunOpentextSection.class);

    @Autowired
    public AutomationClick automationClick;

    @Value("${file.batch.config}")
    private String batchConfigLocation;

    @Value("${file.web.driver.chrome.path}")
    private String fileChromeDriver;

    @Value("${max.waiting.time}")
    private int maxOpentextWaitingTime;

    private ChromeDriver driver;
    //private WebDriverWait wait;

    private boolean testSuccess = false;
    private String password;
    private String errorMessage;
    private String templateName;

    public RunOpentextSection() {

    }

    // Setup web driver
    public void initiateChromeDriver() throws TestFailedException {
        try {
            Path pathToChromeDriver = Paths.get(fileChromeDriver);

            if(!Files.exists(pathToChromeDriver)){
                errorMessage = "File chrome driver was not available at " + pathToChromeDriver;
                informTestFailed(errorMessage);
            }

            System.setProperty("java.awt.headless", "false");
            System.setProperty("webdriver.chrome.driver", String.valueOf(pathToChromeDriver));
            // Disable ChromeDriver startup logs
            System.setProperty("webdriver.chrome.silentOutput", "true");

            driver = new ChromeDriver();
            driver.manage().window().maximize();
            //wait = new WebDriverWait(driver, 20); //100
            LOG.info("Chrome driver was being initialized");
        } catch (WebDriverException e){
            e.printStackTrace();
            errorMessage = "Could not initiate web driver. Please check whether Chrome and chromedriver have the same version.";
            informTestFailed(errorMessage);
        }
    }

    // Open the login page
    public void openLoginPage(String loginPage) throws TestFailedException {
        try {
            driver.get(loginPage);
            Thread.sleep(MEDIUM_WAIT);
            LOG.info("Opened URL: " + loginPage);

            // Switch to new Chrome window
            for (String currentWindow: driver.getWindowHandles()){
                driver.switchTo().window(currentWindow);
            }
            Thread.sleep(SHORT_WAIT);
        } catch (InterruptedException e){
            errorMessage = "Problem with Thread handling";
            informTestFailed(errorMessage);
        }
    }

    // Get user's password from the batch config file
    public void getPassword(String fileLocation, String username) {
        try {
            File batchConfig = new File(fileLocation);
            if (!batchConfig.exists()) {
                errorMessage = "Could not find batch config file.\n" +
                        "Please check if the file is at: " + fileLocation;
                informTestFailed(errorMessage);
            }
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(batchConfig);
            doc.getDocumentElement().normalize();

            NodeList nList = doc.getElementsByTagName("USER");

            for (int nListElement = 0; nListElement < nList.getLength(); nListElement++) {
                Node nNode = nList.item(nListElement);

                if (nNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element eElement = (Element) nNode;
                    if (eElement.getElementsByTagName("USERID").item(0).getTextContent().equals(username)) {
                        password = eElement.getElementsByTagName("PASSWORD").item(0).getTextContent();
                        LOG.info("Password is retrieved successfully");
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Input the credentials and click log in
    public void enterCredentials(String username) throws TestFailedException {
        try {
            driver.findElement(By.xpath("//*[@id=\"s_swepi_1\"]")).sendKeys(username);
            Thread.sleep(SHORT_WAIT);
            driver.findElement(By.xpath("//*[@id=\"s_swepi_2\"]")).sendKeys(password);
            driver.findElement(By.xpath("//*[@id=\"s_swepi_2\"]")).sendKeys(Keys.ENTER);
            Thread.sleep(LONG_WAIT);
            LOG.info("Entered the user's credentials");
        } catch (NoSuchElementException e){
            errorMessage = "There were no fields to enter credentials.";
            informTestFailed(errorMessage);
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Could not log in the system. Please check your username or password";
            informTestFailed(errorMessage);
        }
    }

    // Input the feedback ID and search it
    public void inputFeedbackID(String feedbackID) throws TestFailedException {
        try {
            driver.findElement(By.name("s_4_1_21_0")).sendKeys(feedbackID);
            driver.findElement(By.name("s_4_1_21_0")).sendKeys(Keys.ENTER);
            Thread.sleep(MEDIUM_WAIT);
            LOG.info("Entered feedback ID: " + feedbackID);
        } catch (NoSuchElementException e){
            errorMessage = "There were no fields to enter feedback information.";
            informTestFailed(errorMessage);
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Could not fill in the feedback ID: " + feedbackID;
            informTestFailed(errorMessage);
        }
    }

    // Open Opentext window and expand the template list
    public void openOpentextWindow(int option) throws TestFailedException {
        try {
            switch (option){
                // Click on Create Comm. button
                case 0:
                    //driver.findElement(By.id("s_1_1_10_0_Ctrl")).click();
                    driver.findElement(By.xpath("//button[@aria-label='Feedback:Create Comm.']")).click();
                    break;

                // Click on Internal Comm. button
                case 1:
                    //driver.findElement(By.id("s_1_1_119_0_Ctrl")).click(); //previously it was 116
                    driver.findElement(By.xpath("//button[@aria-label='Feedback:Internal Comm.']")).click();
                    break;

                // Click on Export Non-Feedback button
                case 2:
                    //driver.findElement(By.id("s_1_1_119_0_Ctrl")).click(); //previously it was 116
                    driver.findElement(By.xpath("//button[@aria-label='Feedback:Export Non-Feedback']")).click();
                    break;

                default:
                    errorMessage = "Please choose only option 0, 1, or 2";
                    informTestFailed(errorMessage);
                    break;
            }
            Thread.sleep(MEDIUM_WAIT);
            LOG.info("Opened Opentext window");

            // Switch to Opentext window
            for (String currentWindow: driver.getWindowHandles()){
                driver.switchTo().window(currentWindow);
            }
            Thread.sleep(SHORT_WAIT);

            // Click on the button Open Legodo.ProtocolHandler
            clickOnLegodoButton();

            // Click on the first drop down arrow
            List<WebElement> arrowArray =  driver.findElementsByClassName("DPCYUNC-L-h");
            arrowArray.get(0).click();
            Thread.sleep(SHORT_WAIT);
            // Click on the second drop down arrow
            arrowArray = driver.findElementsByClassName("DPCYUNC-L-h");
            arrowArray.get(1).click();
            Thread.sleep(SHORT_WAIT);
            LOG.info("Expanded the templates list");
        } catch (NoSuchElementException e){
            e.printStackTrace();
            errorMessage = "Could not find \"Create\" button.";
            informTestFailed(errorMessage);
        } catch (IndexOutOfBoundsException | ElementNotInteractableException e){
            errorMessage = "Could not click on the arrow button";
            informTestFailed(errorMessage);
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Could not expand the templates list";
            informTestFailed(errorMessage);
        }
    }

    // Select a specific template in the list
    public void selectTemplate(String templateName) throws TestFailedException {
        try {
            this.templateName = templateName;
            // Click on the template
            try {
                WebElement template = driver.findElement(By.xpath("//*[@class='DPCYUNC-L-m' and text()='" + templateName + "']"));
                if (template.isDisplayed()){
                    template.click();
                    Thread.sleep(SHORT_WAIT);
                }
            } catch (NoSuchElementException e){
                errorMessage = "The template did not exist in the current template list.";
                informTestFailed(errorMessage);
            }

            // Click on "Create" button
            driver.findElement(By.id("tcdr-btn-create")).click();
            Thread.sleep(MEDIUM_WAIT);

            // If second "Create" button appears, click it
            if (driver.findElement(By.id("tcd-btn-back")).isDisplayed()){
                driver.findElement(By.id("tcdr-btn-create")).click();
                Thread.sleep(SUPER_WAIT);
            } else {
                Thread.sleep(LONG_WAIT);
            }
            LOG.info("Opened template : " + templateName);
        } catch (NoSuchElementException e){
            LOG.info("This template did not require to click Create twice.");
        } catch (NoSuchWindowException ignored){

        } catch (Exception e){
            //e.printStackTrace();
            errorMessage = "Could not open template: " + templateName;
            informTestFailed(errorMessage);
        }
    }

    // Set focus on Word window. If Word is not opened, test stops.
    public void focusOnWordWindow() throws TestFailedException {
        try {
            int waitingTime = 0;
            HWND hwnd = new HWND();
            // Wait for opening Word
            while (waitingTime <= maxOpentextWaitingTime){
                // HWND hwnd = User32.INSTANCE.FindWindow(null, "Dokument1 - Word"); // window title

                // Check whether Opentext is available
                checkAlertPartnerSystem();

                // Check whether PowerDocs alert message is displayed
                checkAlertPowerDocs();

                // Check whether printer alert message is displayed
                // Usually, a PowerDocs popup appears when a Printer window pops up
                checkAlertPrinter();

                hwnd = User32.INSTANCE.FindWindow("OpusApp", null);
                if (hwnd == null) {
                    LOG.info("Waiting for Word window to open. Time spent: " + (float)(waitingTime/1000) + " seconds");
                    waitingTime += MEDIUM_WAIT;
                    Thread.sleep(MEDIUM_WAIT);
                }
                else{
                    // Focus on Word window
                    // HWND hwnd = User32.INSTANCE.FindWindow("OpusApp", null);
                    User32.INSTANCE.ShowWindow(hwnd, 9 );      // SW_RESTORE
                    User32.INSTANCE.ShowWindow(hwnd, 3 );      // SW_MAXIMIZE
                    User32.INSTANCE.SetForegroundWindow(hwnd);   // bring to front
                    User32.INSTANCE.SetFocus(hwnd);
                    User32.INSTANCE.GetActiveWindow();
                    LOG.info("Word was opened and focused");
                    Thread.sleep(MEDIUM_WAIT);
                    break;
                }
            }

            // If in the end, Word is still not opened -> Test failed.
            if (hwnd == null) {
                errorMessage = "Word did not open. Words might take time too long to show up.";
                informTestFailed(errorMessage);
            }
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    /*  Choose option to create an attachment on feedback
        0: Click on Save button (SA)
        1: Click on Email button (EM)
        2: Click on Fax button (FA)
        3: Click on Print button (LP)
        4: Click on Central print button (CP)*/
    public void createAttachment(int option) throws TestFailedException {
        try{
            automationClick.closeLicenseWord();

            switch (option){
                // Click on Save button
                case 0:
                    automationClick.selectWordOption(KeyEvent.VK_S, KeyEvent.VK_A);
                    LOG.info("Clicked on Save button.");
                    break;

                // Click on Email button
                case 1:
                    automationClick.selectWordOption(KeyEvent.VK_E, KeyEvent.VK_M);
                    LOG.info("Clicked on Email button.");
                    // In case of appearing pop up to enter recipient email, Word cannot be closed.
                    // Therefore, the pop up will be closed and click on Email button again.
                    if (wordIsOpened()){
                        automationClick.clickSendEmail();
                        LOG.info("Clicked on Send button");
                    }
                    break;

                // Click on Fax button
                case 2:
                    automationClick.selectWordOption(KeyEvent.VK_F, KeyEvent.VK_A);
                    LOG.info("Clicked on Fax button.");
                    break;

                // Click on Print button
                case 3:
                    automationClick.selectWordOption(KeyEvent.VK_L, KeyEvent.VK_P);
                    LOG.info("Clicked on Print button.");
                    break;

                // Click on Central Print button
                case 4:
                    automationClick.selectWordOption(KeyEvent.VK_C, KeyEvent.VK_P);
                    LOG.info("Clicked on Central Print button.");
                    break;

                default:
                    errorMessage = "Please select only option from 0 to 4.";
                    informTestFailed(errorMessage);
                    break;
            }
            Thread.sleep(SHORT_WAIT);
            informTestSucessful();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    // Check whether Word is still opened after clicking an option on Opentext
    public boolean wordIsOpened() throws TestFailedException {
        try {
            String line;
            Process p = Runtime.getRuntime().exec(System.getenv("windir") +"\\system32\\"+"tasklist.exe /v");
            BufferedReader input = new BufferedReader(new InputStreamReader(p.getInputStream()));
            while ((line = input.readLine()) != null) {
                // Check if a Word session with name like template name is still opening
                if (line.contains(templateName) || line.contains("PowerDocs")){
                    input.close();
                    p.destroy();
                    return true;
                }
            }
            input.close();
            p.destroy();
        } catch (IOException e) {
            errorMessage = "Cannot check whether was Word still opening or not.";
            informTestFailed(errorMessage);
        }
        return false;
    }

    public void checkWordIsClosed() throws TestFailedException {
        if (wordIsOpened()){
            errorMessage = "Word was still opening. Please check if the create attachment button is enabled.";
            informTestFailed(errorMessage);
        } else {
            informTestSucessful();
        }
    }

    // Click on the OpenLegodo button of the popup when clicking Create button on Opentext window
    public void clickOnLegodoButton() throws TestFailedException {
        try{
            automationClick.clickLegodo();
            Thread.sleep(SHORT_WAIT);
            LOG.info("Clicked on button Legodo");
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Could not click on Legodo button";
            informTestFailed(errorMessage);
        }
    }

    // Check whether Alert message "PowerDocs is not installed. Please contact your system administrator." pops up
    public void checkAlertPowerDocs() throws TestFailedException {
        try {
            //if (driver.findElementByClassName().isDisplayed()){
            if (driver.findElement(By.className("DPCYUNC-g-g")).isDisplayed()){
                LOG.info("Alert PowerDocs popped up ");
                driver.findElement(By.id("message-box-btn-ok")).click();
                LOG.info("Button \"OK\" on PowerDocs alert was clicked");
                Thread.sleep(SHORT_WAIT);
                clickOnCreateAgain();
            }
        } catch (NoSuchElementException e){
            /*LOG.info("Alert PowerDocs did not pop up");*/
        } catch (NoSuchWindowException e){
            /*LOG.info("Alert PowerDocs turned off");*/
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Problem with popup handling for PowerDocs alert";
            informTestFailed(errorMessage);
        }
    }

    // Check whether Printer option appears
    public void checkAlertPrinter() throws TestFailedException {
        try {
            String alertMessage = "Configure printer";
            if (driver.findElement(By.xpath("//*[contains(text(),'" + alertMessage + "')]")).isDisplayed()){
                LOG.info("Alert Printer popped up ");
                // Click on OK button (id("prd-btn-ok")) will freeze the window -> Click Cancel instead
                driver.findElement(By.id("prd-btn-cancel")).click();
                LOG.info("Button \"Cancel\" on Printer popup was clicked");
                Thread.sleep(SHORT_WAIT);
                clickOnCreateAgain();
            }
        } catch (NoSuchElementException e){
            /*LOG.info("Alert Printer did not pop up");*/
        } catch (NoSuchWindowException e){
            /*LOG.info("Alert Printer turned off");*/
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Problem with popup handling for Printer popup";
            informTestFailed(errorMessage);
        }
    }

    // Check whether Alert message "Partner system not available" pops up
    public void checkAlertPartnerSystem() throws TestFailedException {
        try {
            String alertMessage = "Partner system not available";
            if ((driver.findElement(By.className("DPCYUNC-g-g")).isDisplayed()) &&
                 driver.findElement(By.xpath("//*[contains(text(),'" + alertMessage + "')]")).isDisplayed()){
                errorMessage = "Opentext error - Partner system not available";
                informTestFailed(errorMessage);
            }
        } catch (NoSuchElementException e){
            /*LOG.info("Alert PowerDocs did not pop up");*/
        } catch (NoSuchWindowException e){
            /*LOG.info("Alert PowerDocs turned off");*/
        } catch (Exception e){
            e.printStackTrace();
            errorMessage = "Problem with popup handling for Partner System alert";
            informTestFailed(errorMessage);
        }
    }

    // In case of Word is not showed due to PowerDocs/Printer pops up, click on Create button again
    public void clickOnCreateAgain() throws TestFailedException {
        // Check if Word is opened. If not, click on Create button again.
        try {
            HWND hwnd = User32.INSTANCE.FindWindow("OpusApp", null);
            if (hwnd == null) {
                LOG.info("Word window was not opened due to popup.");
                // Click on "Create" button again
                driver.findElement(By.id("tcdr-btn-create")).click();
                LOG.info("Button \"Create\" was clicked again");
                Thread.sleep(LONG_WAIT);
            }
        } catch (ElementClickInterceptedException | InterruptedException e){
            e.printStackTrace();
            errorMessage = "Could not click on button \"Create\" again";
            informTestFailed(errorMessage);
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    public boolean isTestSuccess() {
        return testSuccess;
    }

    public void informTestFailed(String errorMessage) throws TestFailedException {
        LOG.error("Test attempt " + OpenTextAutomation.attempt + " failed");
        LOG.error("REASON: " + errorMessage);
        LOG.info("ChromeDriver stopped");
        automationClick.takeScreenshot();
        driver.quit();
        throw new TestFailedException();
    }

    public void informTestSucessful(){
        testSuccess = true;
        LOG.info("An attachment is created in the feedback.");
        LOG.info("ChromeDriver stopped");
        LOG.info("Test succeeded");
        driver.quit();
    }
}
