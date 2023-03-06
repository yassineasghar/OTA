package com.example.demo.utility;

import com.example.demo.OpenTextAutomation;
import com.example.demo.exception.TestFailedException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

//@Component
public class ReadBrowsingHistory {

    private static final Logger LOG = LoggerFactory.getLogger(ReadBrowsingHistory.class);

    @Value("${file.history.report.path}")
    private String fileHistoryReportLocation;

    @Value("${file.history.program.path}")
    private String fileHistoryProgramLocation;

    @Value("${wanted.url.prefix}")
    private String wantedURLPrefix;

    private String errorMessage;

    public void terminateActiveSessions() throws IOException {
        ProcessBuilder processBuilder = new ProcessBuilder();
        // Terminate all active Word windows
        processBuilder.command("cmd.exe", "/c", "taskkill /f /im winword.exe").start();
        // Terminate all active BrowsingHistoryView windows
        processBuilder.command("cmd.exe", "/c", "taskkill /f /im BrowsingHistoryView.exe").start();
    }

    public void createBrowsingHistory(int timeInterval) throws IOException, InterruptedException, TestFailedException {
        // Check if the program BrowsingHistoryView exists
        Path pathToFileHistoryProgram = Paths.get(fileHistoryProgramLocation);
        if(!Files.exists(pathToFileHistoryProgram)){
            errorMessage = "Program BrowsingHistoryView is not available";
            informTestFailed(errorMessage);
        }

        Path pathToFileHistoryReport = Paths.get(fileHistoryReportLocation);

        // Create a file recording the last "timeInterval" minutes browsing history
        ProcessBuilder processBuilder = new ProcessBuilder();
        processBuilder.command("cmd.exe", "/c",
                        pathToFileHistoryProgram +
                        " /VisitTimeFilterType 5 /VisitTimeFilterValue " +
                        timeInterval + " /shtml " +
                        pathToFileHistoryReport).start();
    }

    public String getLink() throws TestFailedException {
        try {
            String returnURL;

            // Check if the file exists
            Path pathToFileHistory = Paths.get(fileHistoryReportLocation);
            if(!Files.exists(pathToFileHistory)){
                errorMessage = "File browsing history was not available";
                informTestFailed(errorMessage);
                return null;
            } else {
                LOG.info("File browsing history report was created");
            }

            // Get file browsing history
            File browsingHistoryFile = new File(String.valueOf(pathToFileHistory));
            // Parse html file to Jsoup parser
            Document doc = Jsoup.parse(browsingHistoryFile, "utf-8");
            //Read all the node "a" from html file
            Elements URLs = doc.select("a");

            //Extract the href attribute of each node "a"
            for (Element URL : URLs){
                String linkHref = URL.attr("href");
                //Extract the LATEST link having the wanted prefix
                if (linkHref.contains(wantedURLPrefix)){
                    LOG.info("Found URL: " + linkHref);
                    returnURL = linkHref;
                    return returnURL;
                }
            }
        } catch (Exception e){
            e.printStackTrace();
        }
        errorMessage = "There was no URL contains: " + wantedURLPrefix;
        informTestFailed(errorMessage);
        return null;
    }

    public String extractFeedbackRowID(String linkToFeedbackRowID){
        String feedbackRowID = null;

        String[] URLParts = linkToFeedbackRowID.split("&");
        for (String part : URLParts){
            if (part.contains(wantedURLPrefix)){
                feedbackRowID = part.substring(part.lastIndexOf("=") + 1);
            }
        }
        LOG.info("Extracted feedback row ID was: " + feedbackRowID);
        return feedbackRowID;
    }

    public void informTestFailed(String errorMessage) throws TestFailedException {
        LOG.error("Test attempt " + OpenTextAutomation.attempt + " failed");
        LOG.error("Reason: " + errorMessage);
        throw new TestFailedException();
    }

//    // Get Feedback ID using ReadBrowsingHistory class
//    readBrowsingHistory = new ReadBrowsingHistory();
//    readBrowsingHistory.terminateActiveSessions();
//    readBrowsingHistory.createBrowsingHistory(20);
//    String linkContainsFeedbackRowID = readBrowsingHistory.getLink();
//    String feedbackRowID = readBrowsingHistory.extractFeedbackRowID(linkContainsFeedbackRowID);
}


