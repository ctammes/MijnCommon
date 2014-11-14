package nl.ctammes.common;

import junit.framework.TestCase;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.io.FilenameFilter;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Calendar;
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

//        System.out.println(ini.lees("sectie", "item2"));
        assertEquals("waarde2", ini.lees("sectie", "item2"));
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
        uren.schrijfTijdCel(61,2,10);
        uren.schrijfTijdCel(-1,2,10);
        uren.schrijfWerkboek();
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
        assertEquals("24:12", false, Excel.isTijdCorrect("24:12"));
        assertEquals("20:74", false, Excel.isTijdCorrect("20:74"));
        assertEquals("20:14", true, Excel.isTijdCorrect("20:14"));
        assertEquals("20-14", false, Excel.isTijdCorrect("20-14"));
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
    public void testBerekenTijdverschil() {
        assertEquals("1", Excel.berekenTijdverschil("12:30", "14:00"), "01:30");
        assertEquals("1", Excel.berekenTijdverschil("12:30", "1400"), "");

    }

}
