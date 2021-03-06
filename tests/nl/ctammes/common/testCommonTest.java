package nl.ctammes.common;

import junit.framework.TestCase;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.junit.Test;

import java.io.File;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 26-3-13
 * Time: 20:20
 * To change this template use File | Settings | File Templates.
 */
public class testCommonTest extends TestCase{
    private String projDir = "";
    private String urenlogDir = "";

    @Override
    public void setUp() throws Exception {
        super.setUp();
        if (Diversen.isDesktop()) {
            projDir = "/home/chris/IdeaProjects2/";
        } else {
            projDir = "/home/chris/IdeaProjects/";
        }
        urenlogDir = projDir + "java/Urenlog";
        if (! Diversen.bestaatPad(urenlogDir)) {
            urenlogDir = projDir + "java/urenlog";
        }

    }

    @Test
    public void testIni() {
        MijnIni ini = new MijnIni("test.ini");
        ini.schrijf("sectie", "item1", "waarde1");
        ini.schrijf("sectie", "item2", "waarde2");
        ini.schrijf("sectie", "item3", "123,456");

//        System.out.println(ini.lees("sectie", "item2"));
        assertEquals("waarde2", ini.lees("sectie", "item2"));
        assertEquals("123,456", ini.lees("sectie", "item3", "200,200"));
        assertEquals("waardeX", ini.lees("sectie", "item9", "waardeX"));
    }

    @Test
    public void testIp() {
//        System.out.println("host: " + Diversen.hostnaam());
//        System.out.println("user: " + Diversen.usernaam());
        assertEquals("chris-HP-Compaq-dc7900-Convertible-Minitower", Diversen.hostnaam());
        assertEquals("chris", Diversen.usernaam());
    }

    @Test
    public void testPwd() {
//        System.out.println("pwd: " + Diversen.pwd());
        assertEquals("/media/home_12/chris/IdeaProjects/java/MijnCommon", Diversen.pwd());
    }

    @Test
    public void testCalendar() {
        Calendar cal = Calendar.getInstance();
        System.out.println(cal.get(Calendar.WEEK_OF_YEAR));
    }

    @Test
    public void testWeeknr() {
        System.out.println(Diversen.weekNummer());
        System.out.println(Diversen.weekNummer("27-10-2014"));
    }

    @Test
    public void testWeekdagnr() {
        System.out.println(Diversen.weekdagNummer());
        System.out.println(Diversen.weekdagNummer("31-10-2014"));
        System.out.println(Diversen.weekdagNummer("01-11-2014"));   // zaterdag
        System.out.println(Diversen.weekdagNummer("02-11-2014"));   // zondag
    }

    @Test
    public void testStringToDate() {
        System.out.println(Diversen.stringToDate("31-10-2014"));
        System.out.println(Diversen.stringToDate("31-09-2014"));
        System.out.println(Diversen.stringToDate("31092014"));
        System.out.println(Diversen.stringToDate(""));
    }

    @Test
    public void testVandaag() {
        System.out.println(Diversen.vandaag());
        System.out.println(Diversen.vandaag("yyyy-MM-dd"));
        System.out.println(Diversen.vandaag("EEEE dd-MM-yyyy"));
    }


    @Test
    public void testDatum() {
        System.out.println(Diversen.vandaag());
        System.out.println(Diversen.datumFormat(Diversen.vandaag(), "yyyy-MM-dd"));
        System.out.println(Diversen.datumFormat(Diversen.vandaag(), "EEEE dd-MM-yyyy"));
    }

    @Test
    public void testFileWeekNr() {
        String URENMASK = "cts(\\d{2})\\.xls";
        String fileName = "cts02.xls";
        Matcher mat = Pattern.compile(URENMASK, Pattern.CASE_INSENSITIVE).matcher(fileName);
        if (mat.find()) {
            System.out.println(mat.groupCount());
            System.out.println(Integer.valueOf(mat.group(1)));
        }
    }

    @Test
    public void testSqlite() {

        Sqlite s = new Sqlite("/home/chris/IdeaProjects/java/QueryDb", "QueryDb.db");
        s.openDb();
        System.out.printf("max: %s\n", s.getMax("query", "id"));
        String sql = "insert into query " +
                "(categorie, titel, tekst) " +
                "values('test_max', 'test_max', 'test last_insert_id');";
        try {
            s.executeNoResult(sql);
            sql = "select last_insert_rowid() last_id from query limit 1";
            ResultSet rs = s.execute(sql);
            System.out.println(rs.getInt("last_id"));
            sql = "delete from query where categorie = 'test_max';";
            s.executeNoResult(sql);
        } catch(Exception e) {
            System.out.println(e.getMessage());
        }
    }

    @Test
    public void testExcel()  {
        Excel excel = new Excel("./src/nl/ctammes/common/tests/", "Urenregistratie CT 25.xls");
//        System.out.println(System.getProperty("user.dir"));
        assertEquals("Urenregistratie CT 25.xls", excel.getSheetFile());

        ArrayList<String> temp = excel.getWerkbladen();
        assertEquals("Blad2", temp.get(1));

        assertEquals(true, excel.bestaatWerkblad("Blad1"));
        assertEquals(false, excel.bestaatWerkblad("Blad9"));

    }

    @Test
    public void testExcel2() {
        Excel uren = new Excel(projDir + "uren2013", "CTS45.xls");
        System.out.println(uren.bestaatWerkblad(0));
    }

    @Test
    public void testSplitsPad() {
        String fullpath = projDir + "uren2013/CTS47.xls";
        String[] result = Diversen.splitsPad(fullpath);
        assertEquals("/home/chris/Ideaprojects2/uren2013", result[0]);
        assertEquals("CTS47.xls", result[1]);
    }

    @Test
    public void testCreateNew() {
        String path = urenlogDir;
        String file = "CTS99.xls";
        Excel nieuw = new Excel(path, file);
        System.out.println(nieuw.getWerkbladen());
        assertEquals("Blad1", nieuw.getWerkbladen().get(0));
        nieuw.sluitWerkboek();
        File f = new File(path + File.separatorChar + file);
        f.delete();
    }

    @Test
    public void testSchrijfCel() {
        Excel uren = new Excel(projDir + "uren2013", "CTS47.xls");

        try {
            File oud = new File(projDir + "uren2013/CTS47.xls");
            File nieuw = new File(projDir + "uren2013/CTS99.xls");
            FileUtils.copyFile(oud, nieuw);

            String path = urenlogDir;
            String file = "CTS99.xls";
            Excel excel = new Excel(path, file);
            excel.schrijfTijdCel(0, 0, 123);
//            excel.schrijfCel(2, 1, "Week: 99");

            excel.wisCellen(18, 2, 5);

            excel.schrijfWerkboek();
            assertEquals("", excel.leesCel(18,2));
            excel.sluitWerkboek();

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    @Test
    public void testDatumUitWeeknr() throws Exception {
        String[] dagen = Diversen.weekDatums(44, 2014);
        assertEquals("begin", "27-10-2014", dagen[0]);
        assertEquals("einde", "02-11-2014", dagen[1]);
        assertEquals("1", dagen[0], "27-10-2014");
        assertEquals("2", dagen[1], "02-11-2014");
    }

    @Test
    public void testGetDatumUitWeekDag() {
        assertEquals("27-10-2014", Diversen.datumUitWeekDag(44, 2, 2014));
    }

    @Test
    public void testMinutenNaarDecimaleTekst() {
        assertEquals("0,8", Excel.minutenNaarDecimaleTekst(45));
        assertEquals("1,5", Excel.minutenNaarDecimaleTekst(90));

    }

    @Test
    public void testVorigeWeekNummer() {
        System.out.println(Diversen.vorigeWeekNummer());
        assertEquals("03-01-2014", 1, Diversen.vorigeWeekNummer("03-01-2014"));
        assertEquals("04-01-2010", 53, Diversen.vorigeWeekNummer("04-01-2010"));
        assertEquals("29-12-2009", 52, Diversen.vorigeWeekNummer("29-12-2009"));
        assertEquals("01-01-2009", 1, Diversen.vorigeWeekNummer("01-01-2009"));
    }

    @Test
    public void testWeekdagNaamKort() {
        assertEquals("do", Diversen.weekdagNaamKort("06-11-2014"));
        assertEquals("zo", Diversen.weekdagNaamKort("02-11-2014"));
        assertEquals("zaterdag", Diversen.weekdagNaamLang("08-11-2014"));
    }

    // deze is voor de desktop
    @Test
    public void testVulTijd() {
        Excel uren = new Excel(urenlogDir, "CTS45.xls");
        uren.schrijfTijdCel(6,5,10);

        // lege tijdcel
        String duur = uren.leesCel(38,4);
        System.out.println("leesCel(38,4): " + duur);

        // in/uit tijden
        duur = uren.leesCel(51,2);
        System.out.println("leesCel(51,2):  " + duur);

        int tijd = Excel.tekstNaarTijd("07:45");
        System.out.println("tijd: " + tijd);
        uren.schrijfTijdCellen(51, 2, 5, tijd);
        uren.wisCellen(51, 4, 1);

        uren.schrijfTijdCel(52, 2, Excel.tekstNaarTijd("17:30"));
        uren.schrijfWerkboek();
        duur = uren.leesCel(52,2);
        System.out.println("leesCel(52,2): " + duur);


    }

    public void testSchrijfTijdCel() {
        Excel uren = new Excel(urenlogDir, "CTS45.xls");
        uren.schrijfTijdCel(26,2,10);
        uren.schrijfWerkboek();
        assertEquals("10", "10.0", uren.leesCel(26,2));

        uren.schrijfTijdCel(26,2,70);
        uren.schrijfWerkboek();
        assertEquals("70", "70.0", uren.leesCel(26,2));

        uren.schrijfTijdCel(-1,2,10);
        uren.schrijfWerkboek();
        assertEquals("fout", "70.0", uren.leesCel(26,2)); // ongeldige rij, dus vorige waarde!

        uren.schrijfTijdCel(26,2,0);
        uren.schrijfWerkboek();
        assertEquals("leeg", "", uren.leesCel(26,2));
        uren.sluitWerkboek();

    }

    // deze is voor de laptop
    @Test
    public void testVultijd1() {
        // celwaarde * 24 * 60 = minuten
        // minuten / 60 / 24 = celwaarde
        Excel uren = new Excel(urenlogDir, "CTS45.xls");
        String duur = uren.leesCel(6,2);
        System.out.println(duur);
        duur = uren.leesCel(9,4);
        System.out.println(duur);
        uren.schrijfTijdCel(9,3,21);
        uren.schrijfWerkboek();
        duur = uren.leesCel(9,3);
        System.out.println(duur);
        uren.sluitWerkboek();
    }

    @Test
    public void testVulTijd2() {
        Excel uren = new Excel(urenlogDir, "CTS45.xls");
        String duur = uren.leesCel(61,2);
        System.out.println(duur);

        uren.schrijfTijdCel(62, 2, Excel.tekstNaarTijd("17:30"));
        uren.schrijfWerkboek();
        duur = uren.leesCel(62,2);
        System.out.println(duur);
        uren.sluitWerkboek();

    }

    @Test
    public void testTekstNaarTijd() {
        assertEquals("1050", Excel.tekstNaarTijd("17:30"), 1050);
        assertEquals("0", Excel.tekstNaarTijd(""),0);
    }

    @Test
    public void testHostnaam() {
        String hostnaam = Diversen.hostnaam();
        if (hostnaam.contains("dc7900")) {
            System.out.println(hostnaam + " is desktop");
        } else {
            System.out.println(hostnaam + " is laptop");
        }

    }

    @Test
    public void testBestaatDirFile() {

        assertEquals("1", Diversen.bestaatPad(projDir), true);
        assertEquals("2", Diversen.bestaatPad(projDir + "xxx"), false);
        assertEquals("3", Diversen.bestaatPad(urenlogDir + "/CTS45.xls"), true);
        assertEquals("4", Diversen.bestaatPad(urenlogDir + "/CTS66.xls"), false);
    }

    @Test
    public void testIsTijdCorrect() {
        assertEquals("12:30", Excel.formatTijd("12:30"), "12:30");
        assertEquals("2:30", Excel.formatTijd("2:30"), "02:30");
        assertEquals(":30", Excel.formatTijd(":30"), "00:30");
        assertEquals("1:3", Excel.formatTijd("1:3"), "01:03");
        assertEquals("1:", Excel.formatTijd("1:"), "01:00");

        assertEquals("24:12", false, Excel.isTijdCorrect("24:12"));
        assertEquals("20:74", false, Excel.isTijdCorrect("20:74"));
        assertEquals("20:14", true, Excel.isTijdCorrect("20:14"));
        assertEquals("20-14", false, Excel.isTijdCorrect("20-14"));
        assertEquals("0:30", true, Excel.isTijdCorrect("0:30"));
        assertEquals(":30", true, Excel.isTijdCorrect(":30"));
        assertEquals("leeg", false, Excel.isTijdCorrect(""));

    }

    @Test
    public void testLeesFilenamen() {
        // let op: mask in lowercase!!
        String[] files = Diversen.leesFileNamen(projDir + "/uren2013", "(.*)(\\d{2})(\\.xls)");
        assertTrue(files.length > 0);

    }

    @Test
    public void testDir() {
        assertEquals("/home/chris/IdeaProjects/java/MijnCommon", Diversen.pwd());
    }

    @Test
    public void testBerekenTijdVerschil() {
        assertEquals("1", Excel.berekenTijdVerschil("12:30", "14:00"), "01:30");
        assertEquals("1", Excel.berekenTijdVerschil("8:30", "10:00"), "01:30");
        assertEquals("2", Excel.berekenTijdVerschil("12:30", "1400"), "");
        assertEquals("3", Excel.berekenTijdVerschil("12:30", "1400"), "");

    }

    @Test
    public void testBerekenSom() {
        assertEquals("1", Excel.berekenTijdSom("02:30", "00:45"), "03:15");
        assertEquals("2", Excel.berekenTijdSom("2:30", "0:45"), "03:15");
        assertEquals("3", Excel.berekenTijdSom("12:30", "1400"), "");
        assertEquals("4", Excel.berekenTijdSom("01:39", "00:24"), "02:03");

    }

    @Test
    public void testCellType() {
        // http://apache-poi.1045710.n5.nabble.com/HSSF-formula-cells-not-calculating-td2297567.html
        Excel uren = new Excel(urenlogDir, "CTS45.xls");

        CreationHelper createHelper = uren.getWerkboek().getCreationHelper();
        CellStyle cellStyle = uren.getWerkboek().createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm"));

        int rij = 39;
        HSSFRow row=uren.getWerkblad().getRow(rij);
        if (row != null) {
            for (int kolom = 0; kolom < row.getLastCellNum(); kolom++) {
                Cell cell = uren.getWerkblad().getRow(rij).getCell(kolom);
//                cell.setCellStyle(cellStyle);
                CellStyle x = cell.getCellStyle();
                System.out.printf("%d: %d %s\n", kolom, cell.getCellType(), cell.getCellStyle().getDataFormat());
            }

            rij = 38;
            row=uren.getWerkblad().getRow(rij);
            if (row != null) {
                for (int kolom = 0; kolom < row.getLastCellNum(); kolom++) {
                    Cell cell = uren.getWerkblad().getRow(rij).getCell(kolom);
                    cell.setCellStyle(cellStyle);
                }
            }

        }

    }

    @Test
    public void testRecalculate() {
        Excel uren = new Excel(urenlogDir, "CTS45.xls");

        HSSFWorkbook wb = uren.getWerkboek();
        HSSFSheet sheet = uren.getWerkblad();

        // lege cel vullen met nieuwe waarde
        double tijd = sheet.getRow(32).getCell(3).getNumericCellValue();
        HSSFRow row = sheet.getRow(38);
        Cell cell = row.getCell(3);
        if (cell == null) {
            row.createCell(3, HSSFCell.CELL_TYPE_NUMERIC);
        }
//        cell.setCellValue(tijd);
        cell.setCellValue(Excel.minutenNaarNummer(43));

        // alle formules herberekenen
        HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(sheet, wb);
        for (Iterator rit = sheet.rowIterator(); rit.hasNext();)
        {
            row = (HSSFRow) rit.next();
//            evaluator.setCurrentRow(row);
            for (Iterator cit = row.cellIterator(); cit.hasNext();)
            {
                cell = (HSSFCell) cit.next();
                if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA)
                {
                    String formula = cell.getCellFormula();
                    if (formula != null)
                    {
                        System.out.printf("%d %d %s\n", row.getRowNum(), cell.getColumnIndex(), formula);
                        evaluator.evaluateFormulaCell(cell);
                        cell.setCellFormula (formula); // ADD THIS OR IT WON'T RECALC
                    }
                }
            }
        }
        uren.schrijfWerkboek();
        uren.sluitWerkboek();
    }

    @Test
    public void testRecalculate_1() {

        Excel uren = new Excel(urenlogDir, "CTS45.xls");

        HSSFWorkbook wb = uren.getWerkboek();
        HSSFSheet sheet = uren.getWerkblad();

        // lege cel vullen met nieuwe waarde
        double tijd = sheet.getRow(32).getCell(3).getNumericCellValue();
        HSSFRow row = sheet.getRow(38);
        Cell cell = row.getCell(3);
        if (cell == null) {
            row.createCell(3, HSSFCell.CELL_TYPE_NUMERIC);
        }
//        cell.setCellValue(tijd);
        cell.setCellValue(Excel.minutenNaarNummer(47));

        uren.herberekenWerkblad();

        uren.schrijfWerkboek();
        uren.sluitWerkboek();

    }



}
