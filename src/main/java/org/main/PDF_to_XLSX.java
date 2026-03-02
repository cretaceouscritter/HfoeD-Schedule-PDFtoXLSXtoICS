package org.main;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
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

        	//String filepath = "C:\\Users\\Aiden.Hintersberger\\Downloads\\Stundenplaene_VI_2023_2026_Maerz_-_April_2026.pdf";

            scanner.nextLine();
            System.out.println("Enter Path to pdf File: (e.g. C:/Users/Name/Desktop/Stundenplaene_August.pdf) ");
            String filepath = scanner.nextLine();
            filepath.replace("\\", "/").replace("\"", "");


            File file = new File(filepath);
            PDDocument document = PDDocument.load(file);

            // Define coordinates for the first character in each cell
            float[] xCoordinates = {131.2499F, 270.00018F, 408.7505F, 547.5008F, 686.2511F};
            ArrayList<Float>[] yCoordinates = PDFTextStripper.getCoords(file);
            for(int i = 0; i<yCoordinates.length; i++) {
            	 Collections.sort(yCoordinates[i]);
            }
            
            // 2D-koordinaten-Array erstellen und ausgaben
            Tuple[][] coordinatesArrayTest = create2DArray(xCoordinates, convertToFloatArray(yCoordinates[0]));   
            //print2DArray(coordinatesArrayTest);
            //SwingUtilities.invokeLater(() -> new TupleArrayGUI(coordinatesArrayTest));
            String[][][][] allStuff = new String[gruppenAnzahl][wochenAnzahl][5][14];
            
            // Parse each page
            int gruppe=1;
            int woche=0;
            Tuple[][] coordinatesArray;
            System.out.println(document.getNumberOfPages() + " Seiten eingelesen");
            for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
            	coordinatesArray = create2DArray(xCoordinates, convertToFloatArray(yCoordinates[pageIndex])); 
            	//System.out.println((yCoordinates[pageIndex]).toArray().length);
            	PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                stripper.setSortByPosition(true);

                // Extract text within specified coordinates
                for (int i = 0; i < coordinatesArray.length; i++) {
                	//System.out.print("DAY: " + (i+1) + "\n");
                    for (int j = 0; j < 14; j++) {
                    	Tuple cell = coordinatesArray[i][j];
                    	float x = cell.getX();
                    	float y = cell.getY();
                    	
                    	Tuple cellBelow;
                    	if(j ==coordinatesArray[0].length-1) cellBelow = new Tuple(y,coordinatesArray[i][j].getX() + 200);
                    	else cellBelow = coordinatesArray[i][j+1];
                    	float height = cellBelow.getY() - y; // Kein zusätzliches 5 hinzufügen

                    	// Du kannst die Breite auch dynamisch basierend auf den Koordinaten der Zellen berechnen
                    	Tuple cellRight;
                    	if(i ==coordinatesArray.length-1) cellRight = new Tuple(coordinatesArray[i][j].getX() + 200, 0);
                    	else cellRight = coordinatesArray[i+1][j];
                    	float width = cellRight.getX() - x;

                    	Rectangle rect = new Rectangle((int) x-1, (int) y, (int) width, (int) height);
                    	stripper.setSortByPosition(true);
                    	stripper.addRegion("region" + i, rect);
                    	stripper.extractRegions(document.getPage(pageIndex));

                    	// Extrahiere Text innerhalb der Zelle
                    	String cellText = stripper.getTextForRegion("region" + i);
                    	if(i==1) System.out.println(cellText);
                    	//cellText = 'x'+ cellText+'x';
                    	cellText = cellText.replaceAll("(?<=[A-Za-z\\d\\.])\r", "#");
                    	if(cellText == "\r")cellText ="";
                    	//System.out.println(cellText);
                    	
                        // Print or process the extracted text
                        //System.out.println(stripper.getTextForRegion("region" + i));
                    	
                        allStuff[gruppe-1][woche][i][j] = (j<13)? cellText: allStuff[0][0][0][12];
                    }
                }
                woche++;
                
                if(woche == wochenAnzahl) {
                	System.out.println("Gruppe "+ gruppe + " verarbeitet");
                	woche = 0;
                	gruppe++;
                }
            }
            empty = allStuff[0][0][4][12];
            System.out.println("Empty: [" + empty + "]");
            //print4DArray(allStuff);
            
            //String excelData = toExcel(allStuff);
            //writeExcelFile(excelData, "C:\\Users\\Aiden.Hintersberger\\OneDrive\\Desktop\\Stundenplaene_August.xlsx");
            writeSchedules(allStuff, "C:\\Users\\Aiden.Hintersberger\\OneDrive\\Desktop\\Stundenplaene_Februar.xlsx");
            
            //System.out.println(excelData);
            document.close();
            
            //alles ist aus datei gelesen etc. jetzt datenverarbeitung
            //Stunde[][][][] parsedArray = parseStundenArray(allStuff);
            /*
            // Ausgabe des geparsten Arrays
            for (Stunde[][][] g : parsedArray) {
                for (Stunde[][] w : g) {
                    for (Stunde[] tag : w) {
                    	for(Stunde s: tag) {
                    		System.out.println(s);
                    	}
                    }
                }
            }*/
            //Nikolaustag(parsedArray);
            
            
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    public static Tuple[][] create2DArray(float[] xCoordinates, float[] yCoordinates) {
        Tuple[][] coordinatesArray = new Tuple[xCoordinates.length][yCoordinates.length];
        for (int i = 0; i < xCoordinates.length; i++) {
            for (int j = 0; j < yCoordinates.length; j++) {
                // Verschiebe den Extraktionsbereich nach oben (zum Beispiel um 5 Einheiten)
                float adjustedY = yCoordinates[j] - 10f;
                coordinatesArray[i][j] = new Tuple(xCoordinates[i], adjustedY);
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
    
    public static void print4DArray(String[][][][] array) {
        /*for (int i = 0; i < array.length; i++) {
            System.out.println("gruppe " + i + ":");
            for(int l = 0; l<array[i].length; l++) {
            	System.out.println("woche " + (l+1) + ":");
	            for (int j = 0; j < array[i][l].length; j++) {
	            	System.out.println("Tag: "+ (j+1));
	                for (int k = 0; k < array[i][l][j].length; k++) {
	                    System.out.print("Stunde " + (k+1) + ":\n "+ array[i][l][j][k] + " " + "\n");
	                }
	                System.out.println();
	            }
	            System.out.println();
            }
            System.out.println();
        }*/
    }

    private static int countGroups(String[][][][] array, int woche, int tag, int stunde, int gruppen) {
        int count = 0;
        for (int gruppe = 0; gruppe < gruppen; gruppe++) {
            if (!array[gruppe][woche][tag][stunde].equals(empty)) {
                count++;
            }
        }
        return count;
    }
  
    private static String toExcel(String[][][][] array) {
        StringBuilder excelData = new StringBuilder();

        int gruppen = array.length;
        int wochen = array[0].length;
        int tage = array[0][0].length;
        int stunden = array[0][0][0].length;

        for (int woche = 0; woche < wochen; woche++) {
            excelData.append("Woche ").append(woche + 1).append("\n");
            excelData.append("Stunde,");
            for (int tag = 0; tag < tage; tag++) {
                excelData.append("Tag ").append(tag + 1).append(",");
            }
            excelData.deleteCharAt(excelData.length() - 1); // Remove trailing comma
            excelData.append("\n");

            for (int stunde = 0; stunde < stunden; stunde++) {
                excelData.append(stunde + 1).append(",");
                for (int tag = 0; tag < tage; tag++) {
                    int anzahlGruppen = countGroups(array, woche, tag, stunde, gruppen);
                    excelData.append(anzahlGruppen).append(",");
                }
                excelData.deleteCharAt(excelData.length() - 1); // Remove trailing comma
                excelData.append("\n");
            }
            excelData.append("\n");
        }

        return excelData.toString();
    }

    private static void writeExcelFile(String excelData, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            String[] sheetsData = excelData.split("\n\n");

            // Beispiel für die Abkürzungen der Wochentage
            String[] weekdays = {"Mo", "Di", "Mi", "Do", "Fr"};
         // Beispiel für den Startmontag der ersten Woche
            String startMonday = "04.12.2023"; // DD.MM.YYYY
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
            LocalDate startDate = LocalDate.parse(startMonday, formatter);

            for (int sheetIndex = 0; sheetIndex < sheetsData.length; sheetIndex++) {
            	
                Sheet sheet = workbook.createSheet("Woche " + (sheetIndex + 1));

                String[] rows = sheetsData[sheetIndex].split("\n");
                String[] timeNames = {
                		"",
                		"Stunde",
                        "08:00-08:45",
                        "08:45-09:30",
                        "10:00-10:45",
                        "10:45-11:30",
                        "12:00-12:45",
                        "12:45-13:30",
                        "13:30-14:15",
                        "14:30-15:15",
                        "15:15-16:00",
                        "16:15-17:00",
                        "17:00-17:45",
                        "18:00-18:45",
                        "18:45-19:30",
                        "19:45-20:30"
                    };
                int rownum = 0;

                for (String row : rows) {
                    Row excelRow = sheet.createRow(rownum++);
                    String[] rowData = row.split(",");
                    int cellnum = 0;

                    for (String cellData : rowData) {
                        Cell cell = excelRow.createCell(cellnum++);

                        // Überprüfung, ob der Zellenwert eine Zahl ist
                        try {
                            int value = Integer.parseInt(cellData.trim());
                            cell.setCellValue(value);

                            // Zellenhintergrund basierend auf der Anzahl der Gruppen setzen
                            if (rownum > 1 && cellnum > 1) { // Überspringe die Header-Zeilen und -Spalten
                                CellStyle style = workbook.createCellStyle();
                                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                                // Farbverlauf von Blau zu Rot
                                double ratio = (double) value / 18.0;
                                int red = (int) (255 * ratio);
                                int green = (int) (255 * (1 - ratio));
                                int blue = 0;

                                byte[] rgb = {(byte) red, (byte) green, (byte) blue};
                                style.setFillForegroundColor(new XSSFColor(rgb));

                                cell.setCellStyle(style);
                            }
                        } catch (NumberFormatException e) {
                            // Ignoriere Zellen, die keine Zahl enthalten
                            cell.setCellValue(cellData);
                        }
                    }

                    // Setze den Wochentag in der ersten Zelle jeder Zeile
                    if (rownum == 2) {
                        Cell dayCell;
                        for(int i = 0; i<5; i++) {
                        	dayCell = excelRow.createCell(i+1);
                        	dayCell.setCellValue(weekdays[i]);
                        }
                    }
                 // Setze den Zeitnamen in der ersten Zelle jeder Zeile
                    if (rownum <= timeNames.length) {
                        Cell timeCell = excelRow.createCell(0);
                        timeCell.setCellValue(timeNames[rownum - 1]);
                    }
                    
                    if (rownum == 1) {
                        Cell dateCell = excelRow.createCell(0);
                        dateCell.setCellValue(startDate.format(formatter));
                    }
                    
                 
                }
             // Inkrementiere das Datum für die nächste Woche
                startDate = startDate.plusWeeks(1);
                
            }

            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
                System.out.println("Excel-Tabelle erfolgreich erstellt unter: " + filePath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void print2DArray(Tuple[][] array) {
        for (int i = 0; i < array[0].length; i++) {
            for (int j = 0; j < array.length; j++) {
                System.out.print("(" + array[j][i].getX() + ", " + array[j][i].getY() + "); ");
            }
            System.out.println();
        }
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
                            //if(tag==1)System.out.println(wochenPlan[tag][stunde]);//HIER SCHON ABGESCHNITTEN
                            cell.setCellValue(wochenPlan[tag][stunde].replace("\n",""));
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
    
    public static Stunde[][][][] parseStundenArray(String[][][][] allStuff) {
        Stunde[][][][] parsedArray = new Stunde[allStuff.length][allStuff[0].length][allStuff[0][0].length][allStuff[0][0][0].length];

        for (int gruppe = 0; gruppe < allStuff.length; gruppe++) {
            for (int woche = 0; woche < allStuff[0].length; woche++) {
                for (int tag = 0; tag < allStuff[0][0].length; tag++) {
                    for (int stunde = 0; stunde < allStuff[0][0][0].length; stunde++) {
                        parsedArray[gruppe][woche][tag][stunde] = parseStunde(allStuff[gruppe][woche][tag][stunde]);
                    }
                }
            }
        }

        return parsedArray;
    }

    private static Stunde parseStunde(String stundeString) {
    	if(stundeString==null)return new Stunde();
    	if(!stundeString.equals(empty)) {
	        String[] parts = stundeString.split("\n", 3);
	        Stunde stunde = new Stunde();
	        stunde.setInhalt(parts[0].trim());
	        stunde.setLehrkraft(parts[1].trim());
	        stunde.setRaum(parts[2].trim());
	        stunde.setBelegt();
	
	        return stunde;
    	}else {
    		return new Stunde();
    	}
    }
    
    private static void Nikolaustag(Stunde[][][][] data) {
    	String[] timeNames = {
                "08:00-08:45",
                "08:45-09:30",
                "10:00-10:45",
                "10:45-11:30",
                "12:00-12:45",
                "12:45-13:30",
                "13:30-14:15"
            };
    	ArrayList<String>[] alleRaeume = new ArrayList[7];
    	for(int i = 0; i<7; i++) {
    		alleRaeume[i] = new ArrayList<String>();
    	}
    	for (int i = 0; i < data.length; i++) {
    	    Stunde[][][] gruppe = data[i];
	        Stunde[][] woche = gruppe[0];
	        Stunde[] tag = woche[2];
            for (int l = 0; l < 7; l++) {
            	Stunde stunde = tag[l];
            	if(stunde.getRaum()!=null) {
            		alleRaeume[l].add("[Gruppe " + (i+1) + ": " +stunde.getRaum()+ "]\n");
            	}
            	
            }
    	}
    	for(int i = 0; i<7; i++) {
    		System.out.println("Stunde "+ (i+1) + " " + timeNames[i]);
    		System.out.println(alleRaeume[i].toString().replace(", ", "").replace("[[", "["));
    	}
    }
   
}

