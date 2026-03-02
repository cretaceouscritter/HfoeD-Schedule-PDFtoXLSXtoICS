package keine_ahnung_was_das_is;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Scanner;


public class Main {
	private static Scanner scanner = new Scanner(System.in);
    public static void main(String[] args) {
//    	System.out.println("Enter Path to pdf File: (e.g. C:/Users/Name/Desktop/Stundenplaene_August.pdf) ");
    	String filepath = scanner.nextLine();
    	//String filepath = "C:\\Users\\Aiden.Hintersberger\\OneDrive\\Desktop\\Stundenplaene_August.pdf";
    	filepath.replace("\\", "/").replace("\"", "");
        try {
            System.out.println(filepath);
            // Load the PDF document
            File pdfFile = new File(filepath);
            PDDocument document = PDDocument.load(pdfFile);

            // Create a PDFTextStripper object
            PDFTextStripper customPdfTextStripper = new PDFTextStripper();

            // Extract text from the PDF
            String text = customPdfTextStripper.getText(document);
            document.close();
            //System.out.println(text);

            String[] wochen = splitString(text); //splittet alle wochen und alle gruppen
            String[] gruppen = nachGruppen(wochen);
            

            // Ausgabe der Ergebnisse
            int count = 1;
            for (String part : gruppen) {
            	writeToFile(part,  "gruppe"+ count + ".txt");
            	//System.out.println(part);
            	count++;
            }
           
            writeToFile(text,  filepath.substring(0, filepath.length()-4)+".txt");
            createExcelSheet(text, filepath.substring(0, filepath.length()-4)+".xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

	//METHODS______________________________________________________________________________
    public static void createExcelSheet(String inputString, String outputPath) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Schedule");

        String[] lines = inputString.split("\n");

        int rowNum = 0;
        for (String line : lines) {
            XSSFRow row = sheet.createRow(rowNum++);
            String[] cells = line.split("\\s+", 2);

            XSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(cells[0]);

            XSSFCell cell2 = row.createCell(1);
            if (cells.length > 1) {
                cell2.setCellValue(cells[1]);
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(outputPath)) {
            workbook.write(fileOut);
        }

        workbook.close();
    }
    public static void writeToFile(String content, String filePath) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            if(content != null)writer.write(content);
            System.out.println("String erfolgreich in die Datei geschrieben.");
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Fehler beim Schreiben in die Datei: " + e.getMessage());
        }
    }
    public static String[] splitString(String input) {
        // Verwende den regulären Ausdruck (Regex) "23/\\d{2}" für die Aufteilung
        String[] parts = input.split("23/");
        
        // Füge "23/xx" wieder zu jedem Teil hinzu (außer dem ersten)
        for (int i = 1; i < parts.length; i++) {
            parts[i] = "23/" + parts[i];
        }

        return parts;
    }
    
    private static String[] nachGruppen(String[] wochen) {
    	int anzahlGruppen = 3;
    	System.out.println(anzahlGruppen);
    	String[] gruppen = new String[anzahlGruppen];
    	for(int i = 1; i<wochen.length; i++) {
    		try{
    			gruppen[Integer.valueOf(wochen[i].substring(3,5))-1] += wochen[i];
    		}catch(Exception e) {
    			System.out.println(i+1); //gibt aus wo fehler
    		}
    	}
		return gruppen;
	}
}