package com.whipp.client;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import com.itextpdf.text.Document;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.RectangleReadOnly;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;

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
				currentOutfile = outputFileName + site + ".pptx";
				int numNulls = 0;
				/* Grab the path to each of the site's files */
				ArrayList<String> files = new ArrayList<String>();
				for (int i = 1; i <= numberOfFiles; i++) {
					File folder = new File("file"+i);
					File[] listOfFiles = folder.listFiles();

					boolean broke = false;
					for(File file : listOfFiles){
						if(file.getName().contains(site)){
							files.add(file.getAbsolutePath());
							broke = true;
							break;
						}
					}
					if(!broke){
						files.add(null);
						numNulls ++;
					}
				}

				if(numNulls == numberOfFiles){
					System.out.println("Skipping: " + site + " no files available");
					continue;
				}
				/* Create the output powerpoint */
				Document document = new Document();
				PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream( "output/" + outputFileName + " " + site + ".pdf" ));

				/* Load all the sources into an array */
				PdfReader[] sources = new PdfReader[numberOfFiles];
				for(int i = 0; i < numberOfFiles; i++){
					if(files.get(i) != null){
						System.out.println("Reading: " + files.get(i));
						sources[i] = new PdfReader(files.get(i));
					}else{
						sources[i] = null;
					}
				}
				document.open();
				PdfContentByte cb = writer.getDirectContent();
				PdfImportedPage page;
				for(String command : ordering){
					/* 0-based indeces */
					int fileIdx = Integer.parseInt(command.split(":")[0]) - 1;
					int startSlide = Integer.parseInt(command.split(":")[1].split("-")[0]);
					int endSlide = Integer.parseInt(command.split(":")[1].split("-")[1]);
					if(sources[fileIdx] != null){
						PdfReader pdfToImport = sources[fileIdx];
						for(int i = startSlide; i <= endSlide; i++){
							Rectangle r = pdfToImport.getPageSize(pdfToImport.getPageN(i));
							document.setPageSize(new RectangleReadOnly(r.getWidth(),r.getHeight()));
							document.newPage();
							page = writer.getImportedPage(pdfToImport, i);
							System.out.println(page.getRole().toString()+"=");
							System.out.println(cb.getPdfDocument().isOpen()+"-");
							cb.addTemplate(page, 0, 0);
						}
					}
				}

				/* Creating the file object */
				/* saving the changes to the file */
				System.out.println("Merging done successfully for: " + site);
				if(document.isOpen()){
					document.close();
				}
				for(PdfReader r: sources){
					if(r!=null)
						r.close();
				}
			}
		} catch (Exception e) {
			System.out.println("Merge Failed for: " + currentOutfile);
			e.printStackTrace(); 
		}

		System.exit(0);
	}
}
