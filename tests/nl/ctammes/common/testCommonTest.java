package nl.ctammes.common;

import junit.framework.TestCase;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
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
    @Override
    public void setUp() throws Exception {
        super.setUp();
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

        assertEquals(true, excel.bestaatWerkklad("Blad1"));
        assertEquals(false, excel.bestaatWerkklad("Blad9"));

    }

    @Test
    public void testExcel2() {
        Excel uren = new Excel("/home/chris/Ideaprojects2/uren2013", "CTS47.xls");
        System.out.println(uren.bestaatWerkklad(0));
    }

    @Test
    public void testSplitsPad() {
        String fullpath = "/home/chris/Ideaprojects2/uren2013/CTS47.xls";
        String[] result = Diversen.splitsPad(fullpath);
        assertEquals("/home/chris/Ideaprojects2/uren2013", result[0]);
        assertEquals("CTS47.xls", result[1]);
    }

    @Test
    public void testCreateNew() {
        String path = "/home/chris/Ideaprojects2/uren2013";
        String file = "CTS99.xls";
        Excel nieuw = new Excel(path, file);
        System.out.println(nieuw.getWerkbladen());
        assertEquals("Sheet0", nieuw.getWerkbladen().get(0));
        nieuw.sluitWerkboek();
        File f = new File(path + File.separatorChar + file);
        f.delete();
    }

    @Test
    public void testSchrijfCel() {
        Excel uren = new Excel("/home/chris/Ideaprojects2/uren2013", "CTS47.xls");

        try {
            File oud = new File("/home/chris/Ideaprojects2/uren2013/CTS47.xls");
            File nieuw = new File("/home/chris/Ideaprojects2/uren2013/CTS99.xls");
            FileUtils.copyFile(oud, nieuw);

            String path = "/home/chris/Ideaprojects2/uren2013";
            String file = "CTS99.xls";
            Excel excel = new Excel(path, file);
            excel.schrijfCel(0, 0, 123.0);
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
        System.out.println(dagen[0] + " - " + dagen[1]);

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
    public void testWisCellen() {
        Excel uren = new Excel("/home/chris/Ideaprojects2/java/Urenlog", "CTS45.xls");

    }

    @Test
    public void testVorigeWeekNummer() {
        System.out.println(Diversen.vorigeWeekNummer());
        assertEquals(1, Diversen.vorigeWeekNummer("03-01-2014"));
        assertEquals(53, Diversen.vorigeWeekNummer("04-01-2010"));
        assertEquals(52, Diversen.vorigeWeekNummer("29-12-2009"));
        assertEquals(1, Diversen.vorigeWeekNummer("01-01-2009"));
    }

    @Test
    public void testWeekdagNaamKort() {
        assertEquals("do", Diversen.weekdagNaamKort("06-11-2014"));
        assertEquals("zaterdag", Diversen.weekdagNaamLang("08-11-2014"));
    }

    @Test
    public void testVulTijd() {
        Excel uren = new Excel("/home/chris/IdeaProjects2/java/Urenlog", "CTS45.xls");
        uren.schrijfCel(6,5,10.0);

    }


}