package nl.ctammes.common;

import java.io.File;
import java.io.FilenameFilter;
import java.net.InetAddress;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 23-3-13
 * Time: 14:09
 * To change this template use File | Settings | File Templates.
 */
public class Diversen {

    /**
     * Lees alle filenamen die aan <mask> voldoen uit de opgegeven direcctory
     * (Static, want niet instantie-afhankelijk.)
     * @param dir
     * @param mask  regex file masker
     * @return
     */
    public static String[] leesFileNamen(String dir, String mask) {
        File map = new File(dir);
        final String regex = mask;  // met mask werkt het niet; moet final zijn
        String[] files = map.list(new FilenameFilter() {
            @Override
            public boolean accept(File map, String fileName) {
                return Pattern.matches(regex, fileName.toLowerCase());
            }
        });
        return files;
    }

    /**
     * Geef de hostname
     * @return
     */
    public static String geefHostnaam() {
        try {
            InetAddress addr = InetAddress.getLocalHost();
            byte[] ipAddr = addr.getAddress();
            return addr.getHostName();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Geef de usernaam
     * @return
     */
    public static String geefUsernaam() {
        return System.getProperty("user.name");
    }



}
