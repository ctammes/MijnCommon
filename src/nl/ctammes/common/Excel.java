package nl.ctammes.common;/*

 * Excel.java
 *
 * Created on 29 juni 2007, 8:06
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 *
 * Zie ook: http://poi.apache.org/spreadsheet/quick-guide.html#NewWorkbook
 */

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 *
 * @author TammesC
 */
public class Excel {
    
    private String sheetDirectory = "";     //"C:\Documents and Settings\TammesC\Mijn documenten\Uren"
    private String sheetFile = null;   // xls-bestand
    private String sheetFullName = null;   // volledige naam
    private FileInputStream sheetPath = null;   // xls-bestand incl. path
    private int regelVan = 0;       // eerste dataregel
    private int regelTm = 0;        // laatste dataregel
    private HSSFWorkbook werkboek;  // werkboek
    private HSSFSheet werkblad;     // werkblad

    public Excel() {
    }

    /**
     * Creates a new instance of Excel
     * Create new workbook with one worksheet if it not exists
     */
    public Excel(String xlsDir, String xlsFile) {
        try {
            sheetDirectory = xlsDir;
            sheetFile = xlsFile;
            sheetFullName = xlsDir + File.separatorChar + xlsFile;
            File file = new File(sheetFullName);
            if (! file.exists()) {
                Workbook wb = new HSSFWorkbook();
                FileOutputStream fileOut = new FileOutputStream(sheetFullName);
                wb.createSheet();
                wb.write(fileOut);
                fileOut.close();
            }
            sheetPath = new FileInputStream(file);
            werkboek = new HSSFWorkbook(sheetPath);
            werkblad = werkboek.getSheetAt(0);
        } catch(Exception e) {
            System.out.println(e.getMessage());
        }
    }

    public FileInputStream getSheetPath() {
        return sheetPath;
    }

    public HSSFSheet getWerkblad() {
        return werkblad;
    }

    public HSSFRow getRegel(int regel) {
        return werkblad.getRow(regel);
    }

    public String getSheetDirectory() {
        return sheetDirectory;
    }

    public String getSheetFile() {
        return sheetFile;
    }

    public String getSheetFullName() {
        return sheetFullName;
    }

    public void setSheetDirectory(String sheetDirectory) {
        this.sheetDirectory = sheetDirectory;
    }

    public HSSFWorkbook getWerkboek() {
        return werkboek;
    }

    public void setWerkboek(HSSFWorkbook werkboek) {
        this.werkboek = werkboek;
    }

    /**
     * Sla het spreadsheet bestand op
     */
    public void schrijfWerkboek() {
        try {
            FileOutputStream fileOut = new FileOutputStream(sheetFullName);
            getWerkboek().write(fileOut);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Herberekenen na aanpassen (hoe??)
     */
    public void herberekenWerkboek() {
        try {
            FileOutputStream fileOut = new FileOutputStream(sheetFullName);
            getWerkboek().write(fileOut);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Sluit het spreadsheet bestand
     */
    public void sluitWerkboek() {
        try {
            sheetPath.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Bestaat het werkblad met deze naam
     * @return
     */
    public boolean bestaatWerkblad(String bladnaam) {
        return (werkboek.getSheet(bladnaam) != null);
    }

    /**
     * Bestaat het werkblad met dit nummer
     * @return
     */
    public boolean bestaatWerkblad(int bladnr) {
        return (werkboek.getSheetAt(bladnr) != null);
    }

    /**
     * Geef de namen van de werkbladen van het werkboek
     * @return
     */
    public ArrayList<String> getWerkbladen() {
        ArrayList<String> namen = new ArrayList<String>();
        for (int i = 0; i<werkboek.getNumberOfSheets(); i++) {
            namen.add(werkboek.getSheetName(i));
        }
        return namen;
    }

    /**
     * Opent een werkblad binnen een werkboek
     * @param bladnaam
     */
    public void openWerkblad(String bladnaam) {
        werkblad = werkboek.getSheet(bladnaam);
    }

    /**
     * Geef aantal rijen in het actieve werkblad
     * @return
     */
    public int aantalRijen() {
        return werkblad.getLastRowNum();
    }

    /**
     * Lees numerieke waarde en geef integer terug
     * Het deel achter de decimale punt wordt genegeerd
     * Gebruikt bij uitlezen van tijd
     * @param rij
     * @param kolom
     * @return
     */
    public int leesIntegerCel(int rij, int kolom) {
        String inhoud = leesCel(rij, kolom);
        int waarde = 0;
        if (inhoud != null && ! inhoud.trim().equals("")) {
            waarde = Integer.parseInt(inhoud.split("\\.")[0]);
        }
        return waarde;
    }

    /**
     * Lees een waarde uit een cel en geef die als string terug
     * @param rij rijnummer
     * @param kolom kolomnummer
     * @return de gevonden waarde als string
     */
    public String leesCel(int rij, int kolom) {
        HSSFRow row=werkblad.getRow(rij);
        if (row == null) {
            werkblad.createRow(rij);
        }
        Cell cell = werkblad.getRow(rij).getCell(kolom);
        return celWaarde(cell);
    }

    /**
     * Vul een cel met een tijdwaarde
     * @param rij
     * @param kolom
     * @param waarde
     */
    public void schrijfTijdCel(int rij, int kolom, int waarde) {
        HSSFRow row=werkblad.getRow(rij);
        if (row == null) {
            werkblad.createRow(rij);
            row=werkblad.getRow(rij);
        }
        Cell cell = row.getCell(kolom);
        if (cell == null) {
            row.createCell(kolom, Cell.CELL_TYPE_NUMERIC);
        }
        row.getCell(kolom).setCellValue(minutenNaarNummer(waarde ));
    }

    /**
     * Vul een reeks cel met een identieke tijdwaarde
     * @param rij
     * @param kolom
     * @param waarde
     */
    public void schrijfTijdCellen(int rij, int kolom, int aantal, int waarde) {
        for (int i = 0; i < aantal; i++) {
            schrijfTijdCel(rij, kolom + i, waarde);
        }
    }

    /**
     * Wis een reeks cellen vanaf een startcel
     * @param rij
     * @param kolom
     * @param aantal
     */
    public void wisCellen(int rij, int kolom, int aantal) {
        HSSFRow row=werkblad.getRow(rij);
        if (row != null) {
            int i = 0;
            while (i++ < aantal) {
                row.getCell(kolom++).setCellType(HSSFCell.CELL_TYPE_BLANK);
            }
        }
    }

    /**
     * Vul een cel met een tekst
     * @param rij
     * @param kolom
     * @param waarde
     */
    public void schrijfCel(int rij, int kolom, String waarde) {
        HSSFRow row=werkblad.getRow(rij);
        if (row == null) {
            werkblad.createRow(rij);
        }
        Cell cell = werkblad.getRow(rij).getCell(kolom);
        if (cell == null) {
            row.createCell(kolom, Cell.CELL_TYPE_STRING);
        }
        row.getCell(kolom).setCellValue(waarde);
    }

    /**
     * Vul een reeks cellen met een identieke waarde
     * @param rij
     * @param kolom
     * @param aantal
     * @param waarde
     */
    public void schrijfCellen(int rij, int kolom, int aantal, String waarde) {
        for (int i = 0; i < aantal; i++) {
            schrijfCel(rij, kolom + i, waarde);
        }
    }

    /**
     * Vul een cel met een getal
     * @param rij
     * @param kolom
     * @param waarde
     */
    public void schrijfCel(int rij, int kolom, int waarde) {
        HSSFRow row=werkblad.getRow(rij);
        if (row == null) {
            werkblad.createRow(rij);
        }
        Cell cell = werkblad.getRow(rij).getCell(kolom);
        if (cell == null) {
            row.createCell(kolom, Cell.CELL_TYPE_NUMERIC);
        }
        row.getCell(kolom).setCellValue(waarde);
    }

    /**
     * Vul een reeks cellen met een identieke waarde
     * @param rij
     * @param kolom
     * @param aantal
     * @param waarde
     */
    public void schrijfCellen(int rij, int kolom, int aantal, int waarde) {
        for (int i = 0; i < aantal; i++) {
            schrijfCel(rij, kolom + i, waarde);
        }
    }

    /**
     * Geef de waarde van een cel als tekst terug
     * @param cell
     * @return string
     */
    public String celWaarde(Cell cell) {

        String waarde="";
        if (cell != null) {
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                waarde = Double.toString(nummerNaarMinuten(cell.getNumericCellValue()));
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                waarde = Double.toString(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_STRING:
                waarde = cell.getRichStringCellValue().toString();
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                break;
            default:
                break;
        }
        }
        return waarde;

    }

    /**
     * Omzetten van een tijdsduur in minuten naar een string (hh:mm)
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdWaarde de numerieke celwaarde van een Excel datum/tijd cel
     * @return een string in het formaat hh:mm
     */
    public static String tijdNaarTekst(int tijdWaarde) {
        int uren=0, minuten=0;

        if (tijdWaarde > 0) {
            uren = (int) tijdWaarde / 60;
            minuten = tijdWaarde - uren * 60;
        }

        return String.format("%02d:%02d",uren,minuten);
    }

    /**
     * Omzetten van een tijdsduur in minuten naar een string (hh:mm)
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdWaarde de tijd in minuten
     * @return de tijd in tekst (hh:mm)
     */
    public static String tijdNaarTekst(double tijdWaarde) {
        long uren=0, minuten=0;

        if (tijdWaarde > 0) {
            uren=(long) tijdWaarde / 60;
            minuten =(long) ((tijdWaarde / 60 - uren) * 60);
        }

        return String.format("%02d:%02d",uren,minuten);
    }

    /**
     * Omzetten van een tijdsduur in minuten naar een string (h,m) met de decimale tijd
     * Vb. 45 -> 0,8
     * @param tijdWaarde
     * @return
     */
    public static String minutenNaarDecimaleTekst(int tijdWaarde) {
        return String.format("%.1f", ((tijdWaarde / 60.0) * 100) /100);
    }

    /**
     * Omzetten van een tijdsduur (hh:mm) naar minuten
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdTekst de tijd in tekst (hh:mm)
     * @return de tijd in minuten
     */
    public static int tekstNaarTijd(String tijdTekst) {
        int uren=0, minuten=0;

        if (!tijdTekst.equals("")) {
            String[] tijd=tijdTekst.split(":");

            uren=Integer.valueOf(tijd[0]);
            minuten=Integer.valueOf(tijd[1]);
            return (uren*60)+minuten;
        } else {
            return 0;
        }
    }

    /**
     * Geef de tijd (hh:mm) terug als een uren en minuten deel
     * @param tijdTekst
     * @return
     */
    public static String[] splitsTijd(String tijdTekst) {
        if (!tijdTekst.equals("")) {
            return tijdTekst.split(":");
        } else {
            return null;
        }
    }

    /**
     * Omzetten van een HSSFCell.getNumericCellValue naar minuten
     * @param tijdWaarde de numerieke celwaarde van een Excel datum/tijd cel
     * @return het aantal minuten
     */
    public static int nummerNaarMinuten(double tijdWaarde) {

        int waarde=0;
        if (tijdWaarde>0) {
            waarde=(int) Math.round((tijdWaarde*24)*60);
        }

        return waarde;
    }

    /**
     * Omzetten van minuten naar een HSSFCell.getNumericCellValue
     * @param minuten de tijd in minuten
     * @return de getalwaarde (deel van het etmaal)
     */
    public static double minutenNaarNummer(double minuten) {

        double waarde=0;
        if (minuten>0) {
            waarde=(minuten/60)/24;
        }

        return waarde;
    }


}
