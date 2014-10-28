    //package nl.ctammes.common.tests;

import com.sun.org.apache.xpath.internal.operations.Div;
import junit.framework.TestCase;
import nl.ctammes.common.Diversen;
import nl.ctammes.common.Excel;
import nl.ctammes.common.MijnIni;
import nl.ctammes.common.Sqlite;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FilenameFilter;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Arrays;
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
//        System.out.println("host: " + Diversen.geefHostnaam());
//        System.out.println("user: " + Diversen.geefUsernaam());
        assertEquals("chris-HP-Compaq-dc7900-Convertible-Minitower", Diversen.geefHostnaam());
        assertEquals("chris", Diversen.geefUsernaam());
    }

    @Test
    public void testPwd() {
//        System.out.println("pwd: " + Diversen.geefPwd());
        assertEquals("/media/home_12/chris/IdeaProjects/java/MijnCommon", Diversen.geefPwd());
    }

    @Test
    public void testWeeknr() {
        Calendar cal = Calendar.getInstance();
        System.out.println(cal.get(Calendar.WEEK_OF_YEAR));
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
        Excel uren = new Excel("/home/chris/Ideaprojects2/uren2012", "CTS47.xls");
        System.out.println(uren.bestaatWerkklad(0));
    }
}
