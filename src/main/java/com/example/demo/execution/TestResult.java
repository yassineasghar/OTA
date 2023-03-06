package com.example.demo.execution;
import com.example.demo.exception.TestFailedException;
import com.example.demo.OpenTextAutomation;
import org.openqa.selenium.WebDriver;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

@Component
public class TestResult {

    @Autowired
    private WebDriver driver;

    @Value("${test.success:false}")
    private boolean testSuccess = false;


    public TestResult(){

    }

    public TestResult(WebDriver driver) {
        this.driver = driver;
    }

    private static final Logger LOG = LoggerFactory.getLogger(TestResult.class);

    public void informTestFailed(String errorMessage) throws TestFailedException {
        LOG.error("Test attempt " + OpenTextAutomation.attempt + " failed");
        LOG.error("Reason: " + errorMessage);
        LOG.info("ChromeDriver stopped");
        OpenTextAutomation.attempt++;
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

    public boolean isTestSuccess() {
        return testSuccess;
    }
}
