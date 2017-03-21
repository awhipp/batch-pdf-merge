package com.whipp.client;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class Main {
    //creating empty presentation
    public static void main(String[] args) {

        String currentOutfile = "null.pptx";
        try {
            /* The List of Sites that each file will contain */
            List<String> sites = Files.readAllLines(Paths.get("sites.txt"));
            List<String> ordering = Files.readAllLines(Paths.get("order.txt"));

            /* The dynamic properties that will change */
            InputStream is = new FileInputStream("application.properties");
            Properties properties = new Properties();
            properties.load(is);
            is.close();

            /* The number of files that will be merged */
            int numberOfFiles = Integer.parseInt(properties.getProperty("number.of.files"));
            /* The output file name */
            String outputFileName = properties.getProperty("result.name");

            /* For each site merge the files in the correct order */
            for(String site : sites) {
                currentOutfile = site.replace(" ", "") + "-" + outputFileName + ".pptx";

                /* Grab the path to each of the site's files */
                ArrayList<String> files = new ArrayList<String>();
                for (int i = 1; i <= numberOfFiles; i++) {
                    File folder = new File("file"+i);
                    File[] listOfFiles = folder.listFiles();

                    for(File file : listOfFiles){
                        if(file.getName().contains(site)){
                            files.add(file.getAbsolutePath());
                            break;
                        }
                    }
                }
                /* Ensure that there are the correct files grabbed */
                if(numberOfFiles != files.size()){
                    throw new Exception("Could not find " + numberOfFiles + " files. Only found " + files.size() + ".");
                }

                /* Create the output powerpoint */
                XMLSlideShow ppt = new XMLSlideShow();

                /* Load all the sources into an array */
                XMLSlideShow[] sources = new XMLSlideShow[numberOfFiles];
                for(int i = 0; i < numberOfFiles; i++){
                    FileInputStream inputstream = new FileInputStream(files.get(i));
                    sources[i] = new XMLSlideShow(inputstream);
                    inputstream.close();
                }

                for(String command : ordering){
                    /* 0-based indeces */
                    int fileIdx = Integer.parseInt(command.split(":")[0]) - 1;
                    int startSlide = Integer.parseInt(command.split(":")[1].split("-")[0]) - 1;
                    int endSlide = Integer.parseInt(command.split(":")[1].split("-")[1]) - 1;

                    XSLFSlide[] slidesToImport = sources[fileIdx].getSlides();
                    for(int i = startSlide; i <= endSlide; i++){
                        XSLFSlide slide = ppt.createSlide(slidesToImport[i].getSlideLayout()).importContent(slidesToImport[i]);
                    }

                }

                /* Creating the file object */
                FileOutputStream out = new FileOutputStream(currentOutfile);
                /* saving the changes to the file */
                ppt.write(out);
                System.out.println("Merging done successfully for: " + site);
                out.close();
            }
        } catch (Exception e) {
            System.out.println("Merge Failed for: " + currentOutfile);
            System.out.println(e.getMessage());
        }

        System.exit(0);
    }
}
