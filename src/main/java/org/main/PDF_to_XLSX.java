package org.main;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Scanner;


public class PDF_to_XLSX {
    private static final Scanner scanner = new Scanner(System.in);
    private static String empty;

    public static void main(String[] args) {
        System.out.print("Anzahl Gruppen: ");
        int gruppenAnzahl = scanner.nextInt();

        try {
            /*
            scanner.nextLine();
            System.out.println("Enter Path to pdf File: (e.g. C:/Users/Name/Desktop/Stundenplaene_August.pdf) ");
            String filepath = scanner.nextLine();
            filepath.replace("\\", "/").replace("\"", "");

            File file = new File(filepath);
            PDDocument document = PDDocument.load(file);
            */

            File[] pdfs = new File("INPUT_PDF").listFiles(f -> f.getName().toLowerCase().endsWith(".pdf"));
            if (pdfs == null || pdfs.length == 0) throw new RuntimeException("No PDF in INPUT_PDF");
            File file = pdfs[0];
            PDDocument document = PDDocument.load(pdfs[0]);

            int wochenAnzahl = document.getNumberOfPages()/gruppenAnzahl;
            System.out.printf("Anzahl Wochen: %d\n", wochenAnzahl);

            System.out.println(document.getNumberOfPages() + " Seiten erkannt");

            // Koordinaten der ersten Character von Spalte sind hier hardcoded
            float[] xCoordinates = {131.2499F, 270.00018F, 408.7505F, 547.5008F, 686.2511F};
            // Koordinaten der ersten Character jeder ZEILE werden hier berechnet
            ArrayList<Float>[] yCoordinates = PDFTextStripper.getCoords(file);
            for (int i = 0; i < yCoordinates.length; i++) {
                Collections.sort(yCoordinates[i]);
            }

            // Riesen Array um alle Daten zu speichern: Gruppen x Wochen x Tage x Stunden
            String[][][][] allStuff = new String[gruppenAnzahl][wochenAnzahl][5][14];

            int gruppe = 1;
            int woche = 0;
            Coordinates[][] coordinatesArray;

            // Jede Seite durchgehen
            for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
                //
                coordinatesArray = create2DArray(xCoordinates, convertToFloatArray(yCoordinates[pageIndex]));

                PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                stripper.setSortByPosition(true);

                // Text aus jeder Zelle einzelnd extrahieren
                for (int i = 0; i < 5; i++) {           // Spalten
                    for (int j = 0; j < 14; j++) {      // Zeilen

                        Coordinates cell = coordinatesArray[i][j];
                        float x = cell.X;
                        float y = cell.Y;

                        // Höhe der Zelle
                        Coordinates bottomBorder;
                        if (j == coordinatesArray[0].length - 1) // Falls Zelle ganz unten -> keine weitere drunter
                            bottomBorder = new Coordinates(y, coordinatesArray[i][j].X + 200);
                        else bottomBorder = coordinatesArray[i][j + 1];
                        float height = bottomBorder.Y - y;

                        // Breite der Zelle
                        Coordinates rightBorder;
                        if (i == coordinatesArray.length - 1)
                            rightBorder = new Coordinates(coordinatesArray[i][j].X + 200, 0);
                        else rightBorder = coordinatesArray[i + 1][j];
                        float width = rightBorder.X - x;

                        // Zelle als Rechteck definieren und das Rechteck aus dem Dokument extrahieren
                        Rectangle rect = new Rectangle((int) x - 1, (int) y, (int) width, (int) height);
                        stripper.setSortByPosition(true);
                        stripper.addRegion("region" + i, rect);
                        stripper.extractRegions(document.getPage(pageIndex));

                        // Extrahiere Text innerhalb der Zelle
                        String cellText = stripper.getTextForRegion("region" + i);

                    	cellText = cellText.replaceAll("(?<=[A-Za-z\\d\\.])\r", "#");
                    	if(cellText == "\r")cellText ="";

                        allStuff[gruppe - 1][woche][i][j] = (j < 13) ? cellText : allStuff[0][0][0][12];
                    }
                }
                woche++;

                if (woche == wochenAnzahl) {
                    System.out.println("Gruppe " + gruppe + " verarbeitet");
                    woche = 0;
                    gruppe++;
                }
            }

            writeSchedules(allStuff, "OUTPUT_XLSX/Stundenplaene.xlsx");
            System.out.println("Daten vollständig in 'Stundenplaene.xlsx' geschrieben");
            document.close();



        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static Coordinates[][] create2DArray(float[] xCoordinates, float[] yCoordinates) {
        Coordinates[][] coordinatesArray = new Coordinates[xCoordinates.length][yCoordinates.length];
        for (int i = 0; i < xCoordinates.length; i++) {
            for (int j = 0; j < yCoordinates.length; j++) {
                // Verschiebe den Extraktionsbereich nach oben (manuelle Feinjustierung)
                float adjustedY = yCoordinates[j] - 10f;
                coordinatesArray[i][j] = new Coordinates(xCoordinates[i], adjustedY);
            }
        }
        return coordinatesArray;
    }

    public static float[] convertToFloatArray(ArrayList<Float> yCoordinates) {
        // Create a float array of the same size as the ArrayList
        float[] floatArray = new float[yCoordinates.size()];

        // Copy elements from ArrayList to the array
        for (int i = 0; i < yCoordinates.size(); i++) {
            floatArray[i] = yCoordinates.get(i);
        }

        return floatArray;
    }

    public static void writeSchedules(String[][][][] allStuff, String outputPath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            for (int gruppe = 0; gruppe < allStuff.length; gruppe++) {
                String[][][] gruppenPlan = allStuff[gruppe];
                String gruppenName = "Gruppe" + (gruppe + 1);

                for (int woche = 0; woche < gruppenPlan.length; woche++) {
                    String[][] wochenPlan = gruppenPlan[woche];
                    String wocheName = "Woche" + (woche + 1);

                    Sheet sheet = workbook.createSheet(gruppenName + "_" + wocheName);

                    for (int stunde = 0; stunde < wochenPlan[0].length; stunde++) {
                        Row row = sheet.createRow(stunde);

                        for (int tag = 0; tag < wochenPlan.length; tag++) {
                            Cell cell = row.createCell(tag);
                            cell.setCellValue(wochenPlan[tag][stunde].replace("\n", ""));
                        }
                    }
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream(outputPath)) {
                workbook.write(fileOut);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}