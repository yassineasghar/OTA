package com.example.demo.execution;

import com.example.demo.OpenTextAutomation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import static com.example.demo.execution.TimeConstants.SHORT_WAIT;

@Component
public class AutomationClick {

    private final Robot robot;
    private final Logger LOG = LoggerFactory.getLogger(AutomationClick.class);

    @Value("${key.powerdocs1}")
    private int keyPowerDocs1;

    @Value("${key.powerdocs2}")
    private int keyPowerDocs2;

    @Value("${test.screenshot.path}")
    private String screenshotLocation;

    @Value("${customer.email}")
    private String customerEmail;

    static {
        System.setProperty("java.awt.headless", "false");
    }

    public AutomationClick() throws AWTException {
        robot = new Robot();
        robot.setAutoDelay(SHORT_WAIT);
    }

    public void clickLegodo() {
        // Tick to Always allow Legodo run
        robot.keyPress(KeyEvent.VK_TAB) ;
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_ENTER);

        // Move to the OpenLegodo button
        robot.keyPress(KeyEvent.VK_TAB);
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_ENTER);
    }

    public void selectWordOption(int button1, int button2){
        // Select PowerDocs tab
        // Because on dev and real environment, Opentext has different key combination to select PowerDocs tab
        robot.keyPress(KeyEvent.VK_ALT);
        robot.keyPress(keyPowerDocs1);
        robot.keyPress(keyPowerDocs2);
        robot.keyPress(button1);
        robot.keyPress(button2);

        robot.keyRelease(KeyEvent.VK_ALT);
        robot.keyRelease(keyPowerDocs1);
        robot.keyRelease(keyPowerDocs2);
        robot.keyRelease(button1);
        robot.keyRelease(button2);
    }

    // On the first test run of the batch, there will be a dialog needed to be close.
    public void closeLicenseWord(){
        // Close the "Sign in to set up Office" pop up
        robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        robot.keyPress(KeyEvent.VK_ESCAPE);
        robot.keyRelease(KeyEvent.VK_ESCAPE);

        robot.keyPress(KeyEvent.VK_ESCAPE);
        robot.keyRelease(KeyEvent.VK_ESCAPE);

        // Close the "Accept the license agreement" pop up
        robot.keyPress(KeyEvent.VK_TAB);
        robot.keyPress(KeyEvent.VK_TAB);
        robot.keyPress(KeyEvent.VK_TAB);
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_ENTER);

        robot.keyPress(KeyEvent.VK_ESCAPE);
        robot.keyRelease(KeyEvent.VK_ESCAPE);
    }

    public void clickSendEmail(){
        // Insert email
        StringSelection stringSelection = new StringSelection(customerEmail);
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        clipboard.setContents(stringSelection, stringSelection);

        robot.keyPress(KeyEvent.VK_CONTROL);
        robot.keyPress(KeyEvent.VK_A);
        robot.keyRelease(KeyEvent.VK_CONTROL);
        robot.keyRelease(KeyEvent.VK_A);

        robot.keyPress(KeyEvent.VK_CONTROL);
        robot.keyPress(KeyEvent.VK_V);
        robot.keyRelease(KeyEvent.VK_CONTROL);
        robot.keyRelease(KeyEvent.VK_V);

        // Move to Send button
        for (int i = 0; i < 26; i++) {
            robot.keyPress(KeyEvent.VK_TAB);
        }

        robot.keyPress(KeyEvent.VK_ENTER);
    }

    public void takeScreenshot(){
        try {
            SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
            Date date = new Date();

            String fileFormat = "png";
            String fileName = formatter.format(date) + "_TestAttempt-" + OpenTextAutomation.attempt + "." + fileFormat;
            String fileLocation = screenshotLocation + fileName;

            // Create folder Screenshots if it is not created yet
            File screenshotFolder = new File(screenshotLocation);
            if (!screenshotFolder.exists()){
                screenshotFolder.mkdir();
            }

            Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
            BufferedImage screenFullImage = robot.createScreenCapture(screenRect);
            ImageIO.write(screenFullImage, fileFormat, new File(fileLocation));
            LOG.info("Screenshot is captured at: " + fileLocation);
        } catch (IOException e){
            LOG.error("Cannot take screenshot of failed test.");
        }

    }

}
