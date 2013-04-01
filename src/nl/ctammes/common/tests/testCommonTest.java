package nl.ctammes.common.tests;

import nl.ctammes.common.Diversen;
import nl.ctammes.common.MijnIni;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 26-3-13
 * Time: 20:20
 * To change this template use File | Settings | File Templates.
 */
public class testCommonTest {
    @BeforeClass
    public static void setUp() throws Exception {

    }

    @Test
    public void testIni() {
        MijnIni ini = new MijnIni("test.ini");
        ini.schrijf("sectie", "item1", "waarde1");
        ini.schrijf("sectie", "item2", "waarde2");

        System.out.println(ini.lees("sectie", "item2"));
    }

    @Test
    public void testIp() {
       System.out.println("host: " + Diversen.geefHostnaam());
       System.out.println("user: " + Diversen.geefUsernaam());
    }
}
