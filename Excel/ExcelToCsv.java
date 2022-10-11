/**
* @author = Leonhard Apitz
* @date = start: 05.10.2022  current: 07.10.2022 end: unknown
*/

import java.io.*; // Import von Java I/O
import org.apache.poi.hssf.usermodel.HSSFSheet; // Import von Apache-Bibliotheken für die Interaktion von Java mit Excel Dateien
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.util.Scanner; // Import des Scanners
import java.util.ArrayList; // Import von dynamischen Arrays

public class ExcelToCsv {
	public static void main(String[] args) {
		// Deklaration von Variablen
		File datei;
		String dateipfad;
		
		// Erzeugen eines neuen Dateiauswahldialog-Objektes, Scanners
		JFileChooser chooser = new JFileChooser();
		InputStream input_stream = new FileInputStream;
		Scanner scanner = new Scanner;
		FileWriter file_writer = new FileWriter();
		
		// Auswählen der Excel Datei
		chooser.setDialogTitle("Datei öffnen");
		int rueckgabeAntwort = chooser.showOpenDialog(null);
		if (rueckgabeAntwort == chooser.APPROVE_OPTION) {
			datei = chooser.getSelectedFile();
		}
		
		dateipfad = datei.getAbsolutePath();
		
		if (String.toLowerCase(dateipfad.endsWith(".xlsx"))) {
			
			Workbook workbook = new XSSFWorkbook(datei);
			Sheet sheet;
			try {
				File ausgabe = new File("Aufgabe.csv");
				if (ausgabe.createNewFile()) {
					System.out.println("File created: " + ausgabe.getName());
				} else {
					System.out.println("File already exists.");
				}
			} catch (IOException io_ex) {
				System.out.println(io_ex);
			}
		}
			
			//Lies Zelle und trage sie in die neue Tabelle ein
			for(int i = 0; i < workbook.getNumberOfSheets(); i++) {
				workbook.setActiveSheet(i);
				messwert = String.valueOf(workbook.getSheetAt(i).getRow(27-1).getCell(5-1)); // Indices beginnen mit null!
				
		} else {
			JOptionPane.showMessageDialog(this,
				"Bitte wählen sie nur Excel-Dateien aus.",
				"Warnung",
				JOptionPane.WARNING_MESSAGE);
		}
	}
}
/*
public static void writeDataLineByLine(String filePath)
{
    // first create file object for file placed at location
    // specified by filepath
    File file = new File(filePath);
    try {
        // create FileWriter object with file as parameter
        FileWriter outputfile = new FileWriter(file);
  
        // create CSVWriter object filewriter object as parameter
        CSVWriter writer = new CSVWriter(outputfile);
  
        // adding header to csv
        String[] header = { "Name", "Class", "Marks" };
        writer.writeNext(header);
  
        // add data to csv
        String[] data1 = { "Aman", "10", "620" };
        writer.writeNext(data1);
        String[] data2 = { "Suraj", "10", "630" };
        writer.writeNext(data2);
  
        // closing writer connection
        writer.close();
    }
    catch (IOException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
    }
}
*/