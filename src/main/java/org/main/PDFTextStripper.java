package org.main;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.TextPosition;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedList;

public class PDFTextStripper extends org.apache.pdfbox.text.PDFTextStripper {

    public PDFTextStripper() throws IOException {
        super();
    }
    public LinkedList<String> lines = new LinkedList<String>();
    @Override
    protected void processTextPosition(TextPosition text) {
        super.processTextPosition(text);
        lines.add(text.getUnicode() +"," + text.getXDirAdj() + "," + text.getYDirAdj());
    }

    public static ArrayList<Float>[] getCoords(File file) {
        try {
        	PDDocument document = PDDocument.load(file);
        	//CustomPDFTextStripper customStripper = new CustomPDFTextStripper();

        	// Process each page
        	ArrayList<Float>[] coordinates = new ArrayList[document.getNumberOfPages()];
        	PDFTextStripper[] strippers = new PDFTextStripper[document.getNumberOfPages()];
        	for (int i = 0; i < document.getNumberOfPages(); i++) {
        		
        	    // Create a new yCoordinates list for each page
        	    coordinates[i] = new ArrayList<Float>();
        	    strippers[i] = new PDFTextStripper();
        	    
        	    strippers[i].setStartPage(i +1);
        	    strippers[i].setEndPage(i +1);

        	    // Extract text and print coordinates
        	    strippers[i].getText(document);

        	    String[][] splitted = strippers[i].splitLines( strippers[i].lines);
        	    for (String[] s : splitted) {
        	        // System.out.println(s[0] + ": " + s[1] + ", " + s[2]);
        	        if (s[0].equals(":")&& Float.valueOf(s[1])<131F) {
        	            coordinates[i].add(Float.valueOf(s[2]));
        	        }
        	    }

        	    // Doppelte entfernen
        	    coordinates[i] = new ArrayList<>(new HashSet<>(coordinates[i]));
        	}

        	document.close();
        	return coordinates;

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
       
    }
    private String[][] splitLines(LinkedList<String> lines) {
        String[][] resultArray = new String[lines.size()][3];
        int i = 0;
        for (String line : lines) {
            // Bei "," splitten und die Teile in das zweidimensionale Array einfügen
            String[] parts = line.split(",");
            resultArray[i]= parts;
            i++;
        }

        return resultArray;
    }
}


