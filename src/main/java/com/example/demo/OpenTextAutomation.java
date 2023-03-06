package com.example.demo;

import com.example.demo.exception.TestFailedException;
import com.example.demo.execution.RunOpentextSection;
//import com.example.demo.execution.TestResult;
import com.example.demo.utility.FeedbackExtractor;
//import com.example.demo.utility.ReadBrowsingHistory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
@SpringBootApplication(exclude = {OpenTextAutomation.class})
//@SpringBootApplication
public class OpenTextAutomation implements CommandLineRunner {

    @Autowired
    private RunOpentextSection opentextSection;

    @Autowired
    private FeedbackExtractor feedbackExtractor;

    //@Autowired
    //private TestResult testResult;

    @Value("${login.page}")
    private String loginPage;

    @Value("${file.batch.config}")
    private String batchConfigFile;

    @Value("${user.username}")
    private String username;

    @Value("${test.report.path}")
    private String testReportLocation;

    @Value("${max.attempts}")
    private int maxAttempts;

    @Value("${run.option}")
    private int runOption;

    @Value("${template.name}")
    private String templateName;

    @Value("${attachment.option}")
    private int createAttachmentOption;

    public static int attempt = 1;

    private static final Logger LOG = LoggerFactory.getLogger(OpenTextAutomation.class);

    public static void main(String[] args) {
        SpringApplication.run(OpenTextAutomation.class, args);
    }

    @Override
    public void run(String... args) {
        runOpentext();
}


    // Run Opentext process
    public void runOpentext() {
        while (!opentextSection.isTestSuccess() && (attempt <= maxAttempts)){
            try {
                LOG.info("==================STARTING NEW OPENTEXT SESSION==================");
                LOG.info("Running attempt: " + attempt + ". Maximum attempt: " + maxAttempts);

                // Get Feedback ID
                feedbackExtractor.terminateActiveSessions();
                String feedbackID = feedbackExtractor.logReader(testReportLocation);

                // Find Feedback ID
                opentextSection.initiateChromeDriver();
                opentextSection.openLoginPage(loginPage);
                opentextSection.getPassword(batchConfigFile, username);
                opentextSection.enterCredentials(username);
                opentextSection.inputFeedbackID(feedbackID);

                // Run Opentext process
                opentextSection.openOpentextWindow(runOption);
                opentextSection.selectTemplate(templateName);
                opentextSection.focusOnWordWindow();
                opentextSection.createAttachment(createAttachmentOption);
                //opentextSection.checkWordIsClosed();
            } catch (TestFailedException | IOException e){
                LOG.info("Test attempt " + attempt + " stopped");
                attempt++;
            }
        }
    }
}
