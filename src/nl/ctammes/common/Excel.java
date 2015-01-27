package nl.ctammes.common;/*

 * Excel.java
 *
 * Created on 29 juni 2007, 8:06
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 *
 * Zie ook: http://poi.apache.org/spreadsheet/quick-guide.html
 */

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.logging.Logger;

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

    protected Logger log;

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

    public void setLog(Logger log) {
        this.log = log;
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
            herberekenWerkblad();
            FileOutputStream fileOut = new FileOutputStream(sheetFullName);
            getWerkboek().write(fileOut);
            fileOut.close();
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
     * Controleer of er een cel is op de aangegeven positie; zo niet, maar die dan aan
     * @param rij
     * @param kolom
     * @return HSSFRow of null bij ongeldige rij/kolom
     */
    private HSSFRow prepareCel(int rij, int kolom) {
        if (rij >=0 && kolom >= 0) {
            HSSFRow row=werkblad.getRow(rij);
            if (row == null) {
                werkblad.createRow(rij);
                row=werkblad.getRow(rij);
            }
            Cell cell = werkblad.getRow(rij).getCell(kolom);
            if (cell == null) {
                row.createCell(kolom, HSSFCell.CELL_TYPE_BLANK);
            }
            return row;
        } else {
            // TODO wat doen als invoer fout is?
            return null;
        }
    }

    /**
     * Vul een cel met een tijdwaarde
     * @param rij
     * @param kolom
     * @param waarde (minuten)
     */
    public void schrijfTijdCel(int rij, int kolom, int waarde) {
        HSSFRow row = prepareCel(rij, kolom);
        if (row != null) {
            if (waarde == 0) {
                row.getCell(kolom).setCellType(HSSFCell.CELL_TYPE_BLANK);
            } else {
                row.getCell(kolom).setCellValue(minutenNaarNummer(waarde ));
            }
        }
    }

    /**
     * Vul een reeks cel met een identieke tijdwaarde
     * @param rij
     * @param kolom
     * @param waarde (minuten)
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
        HSSFRow row = prepareCel(rij, kolom);
        if (row != null) {
            int i = 0;
            while (i++ < aantal) {
                row.getCell(kolom++).setCellType(HSSFCell.CELL_TYPE_BLANK);
            }
        }
    }

    /**
     * Wist een rij uit het werkblad
     * @param rij
     */
    public void wisRij(int rij) {
        if (rij >= 0) {
            HSSFRow row = werkblad.getRow(rij);
            werkblad.removeRow(row);
            werkblad.shiftRows(rij + 1, werkblad.getLastRowNum(), -1);
            schrijfWerkboek();
            if (log != null) {
                log.info(String.format("rij %d verwijderd", rij));
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
        HSSFRow row = prepareCel(rij, kolom);
        if (row != null) {
            row.getCell(kolom).setCellValue(waarde);
        }
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
        HSSFRow row = prepareCel(rij, kolom);
        if (waarde == 0) {
            row.getCell(kolom).setCellType(HSSFCell.CELL_TYPE_BLANK);
        } else {
            row.getCell(kolom).setCellValue(waarde);
        }
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
     * Herbereken het werkblad
     * Doorloop alle cellen en herberelen iedere formule
     *
     * http://apache-poi.1045710.n5.nabble.com/HSSF-formula-cells-not-calculating-td2297567.html
     */
    public void herberekenWerkblad() {

        HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(werkblad, werkboek);

        for (Iterator rit = werkblad.rowIterator(); rit.hasNext(); ) {
            HSSFRow row = (HSSFRow) rit.next();
            for (Iterator cit = row.cellIterator(); cit.hasNext(); ) {
                HSSFCell cell = (HSSFCell) cit.next();
                if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
                    String formula = cell.getCellFormula();
                    if (formula != null) {
                        evaluator.evaluateFormulaCell(cell);
                        cell.setCellFormula(formula); // ADD THIS OR IT WON'T RECALC
                    }
                }
            }
        }
//        schrijfWerkboek();
//        sluitWerkboek();
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

        if (uren < 24 && minuten < 60) {
            return String.format("%02d:%02d", uren, minuten);
        } else {
            return "";
        }
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

        if (isTijdCorrect(tijdTekst)) {
            String[] tijd=splitsTijd(tijdTekst);

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
        if (! tijdTekst.equals("") && tijdTekst.contains(":")) {
            return formatTijd(tijdTekst).split(":");
        } else {
            return null;
        }
    }

    /**
     * Formatteer een tijd (hh:mm) inclusief voorloopnullen
     * @param tijdTekst (hh:mm)
     * @return (hh:mm)
     */
    public static String formatTijd(String tijdTekst) {
        String result = "";
        if (! tijdTekst.equals("") && tijdTekst.contains(":")) {
            String[] tijd = tijdTekst.split(":");
            int uur = tijd[0].equals("") ? 0 : Integer.parseInt(tijd[0]);
            int min = 0;
            if (tijd.length == 2) {
                min = tijd[1].equals("") ? 0 : Integer.parseInt(tijd[1]);
            }
            result = String.format("%02d:%02d", uur, min);
        }
        return result;
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

    /**
     * Controleer of de tijd geldig is
     * @param tijdTekst (hh:mm)
     * @return
     */
    public static boolean isTijdCorrect(String tijdTekst) {
        String[] tijd = splitsTijd(tijdTekst);
        if (tijd != null   &&  tijd.length == 2) {
            int uren = Integer.valueOf(tijd[0]);
            int minuten = Integer.valueOf(tijd[1]);
            return (uren < 24 && minuten < 60);
        } else {
            return false;
        }
    }

    /**
     * Berekent het verschil tussen twee tijdstippen
     * @param tijdTekst1 (hh:mm)
     * @param tijdTekst2 (hh:mm)
     * @return (hh:mm) leeg bij foute invoer
     */
    public static String berekenTijdVerschil(String tijdTekst1, String tijdTekst2) {
        int tijd1 = tekstNaarTijd(tijdTekst1);
        int tijd2 = tekstNaarTijd(tijdTekst2);
        String result = "";
        if (tijd1 == 0 || tijd2 == 0) {
            return "";
        } else if (tijd1 > tijd2) {
            result = tijdNaarTekst(tijd1 - tijd2);
        } else {
            result = tijdNaarTekst(tijd2 - tijd1);
        }
        return result;
    }

    /**
     * Berekent de som van twee tijdstippen
     * @param tijdTekst1 (hh:mm)
     * @param tijdTekst2 (hh:mm)
     * @return (hh:mm) leeg bij foute invoer
     */
    public static String berekenTijdSom(String tijdTekst1, String tijdTekst2) {
        int tijd1 = tekstNaarTijd(tijdTekst1);
        int tijd2 = tekstNaarTijd(tijdTekst2);
        String result = "";
        if (tijd1 == 0 || tijd2 == 0) {
            return "";
        }
        result = tijdNaarTekst(tijd1 + tijd2);
        return result;
    }
}
