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
        System.out.print("Anzahl Wochen: ");
        int wochenAnzahl = scanner.nextInt();
        try {

            scanner.nextLine();
            System.out.println("Enter Path to pdf File: (e.g. C:/Users/Name/Desktop/Stundenplaene_August.pdf) ");
            String filepath = scanner.nextLine();
            filepath.replace("\\", "/").replace("\"", "");


            File file = new File(filepath);
            PDDocument document = PDDocument.load(file);

            // Define coordinates for the first character in each cell
            float[] xCoordinates = {131.2499F, 270.00018F, 408.7505F, 547.5008F, 686.2511F};
            ArrayList<Float>[] yCoordinates = PDFTextStripper.getCoords(file);
            for (int i = 0; i < yCoordinates.length; i++) {
                Collections.sort(yCoordinates[i]);
            }

            String[][][][] allStuff = new String[gruppenAnzahl][wochenAnzahl][5][14];

            // Parse each page
            int gruppe = 1;
            int woche = 0;
            Coordinates[][] coordinatesArray;
            System.out.println(document.getNumberOfPages() + " Seiten eingelesen");

            for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
                coordinatesArray = create2DArray(xCoordinates, convertToFloatArray(yCoordinates[pageIndex]));

                PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                stripper.setSortByPosition(true);

                // Extract text within specified coordinates
                for (int i = 0; i < coordinatesArray.length; i++) {

                    for (int j = 0; j < 14; j++) {
                        Coordinates cell = coordinatesArray[i][j];
                        float x = cell.X;
                        float y = cell.Y;

                        Coordinates cellBelow;
                        if (j == coordinatesArray[0].length - 1)
                            cellBelow = new Coordinates(y, coordinatesArray[i][j].X + 200);
                        else cellBelow = coordinatesArray[i][j + 1];
                        float height = cellBelow.Y - y; // Kein zusätzliches 5 hinzufügen

                        // Breite dynamisch basierend anhand der Koordinaten der Zellen berechnen
                        Coordinates cellRight;
                        if (i == coordinatesArray.length - 1)
                            cellRight = new Coordinates(coordinatesArray[i][j].X + 200, 0);
                        else cellRight = coordinatesArray[i + 1][j];
                        float width = cellRight.X - x;

                        Rectangle rect = new Rectangle((int) x - 1, (int) y, (int) width, (int) height);
                        stripper.setSortByPosition(true);
                        stripper.addRegion("region" + i, rect);
                        stripper.extractRegions(document.getPage(pageIndex));

                        // Extrahiere Text innerhalb der Zelle
                        String cellText = stripper.getTextForRegion("region" + i);

                        cellText = cellText.replaceAll("(?<=[A-Za-z\\d\\.])\r", "#");
                        if (cellText == "\r") cellText = "";

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

            writeSchedules(allStuff, "Stundenplaene.xlsx");
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
                // Verschiebe den Extraktionsbereich nach oben (zum Beispiel um 5 Einheiten)
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