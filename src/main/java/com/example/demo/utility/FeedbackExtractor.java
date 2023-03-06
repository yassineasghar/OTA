package com.example.demo.utility;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

@Component
public class FeedbackExtractor {

    private static final Logger LOG = LoggerFactory.getLogger(FeedbackExtractor.class);

    public String logReader(String myDirectoryPath) {
        String feedbackNumber = "";
        File root = new File(myDirectoryPath);
        File[] topDirectories = root.listFiles(new FileFilter() {
            public boolean accept(File file) {
                return file.isDirectory();
            }
        });
        //System.out.println("topDirectories = " + Arrays.toString(topDirectories));

        for (File directory : topDirectories) {
            File recentFile = new File("");
            //recentFile.setLastModified(0L);
            recentFile = getRecentFile(directory, recentFile);
            if(recentFile.getName().contains("html")){
                //System.out.println("recentFile = " + recentFile);
                String text = readWithGoogleGuava(recentFile.getAbsolutePath());
                int feedbackLength = 9;
                int trim = 1;
                int startIndex = text.lastIndexOf("is stored in the variable") - feedbackLength;
                int endIndex = text.lastIndexOf("is stored in the variable") - trim;
                feedbackNumber = text.substring(startIndex, endIndex);
                LOG.info("Feedback number was: " + feedbackNumber);
            }
        }

        return feedbackNumber;
    }

    private File getRecentFile(File dir, File baseFile) {
        File recentFile = baseFile;
        for (File file : dir.listFiles()) {
            if (file.isDirectory()) {
                recentFile = getRecentFile(file, recentFile);
            } else {
                if (file.lastModified() > recentFile.lastModified() && file.getName().contains("1-"))  {
                    recentFile = file;
                }
            }
        }
        return recentFile;
    }

    private String readWithGoogleGuava(String fileToRead) {
        String readData = "";
        try {
            readData = com.google.common.io.Files.asCharSource(new File(fileToRead), StandardCharsets.UTF_8).read();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return readData;
    }

    public void terminateActiveSessions() throws IOException {
        ProcessBuilder processBuilder = new ProcessBuilder();
        // Terminate all active Word windows
        processBuilder.command("cmd.exe", "/c", "taskkill /f /im winword.exe").start();
        LOG.info("Termiate all previous active Word sessions.");
    }
}
