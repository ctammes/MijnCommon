package nl.ctammes.common;

import com.sun.org.apache.xpath.internal.SourceTree;
import org.apache.poi.ss.formula.functions.Match;

import java.io.File;
import java.io.FilenameFilter;
import java.net.InetAddress;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
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
    public static String hostnaam() {
        try {
            InetAddress addr = InetAddress.getLocalHost();
            byte[] ipAddr = addr.getAddress();
            return addr.getHostName();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * geef het huidige weeknummer
     * @return het weeknummer
     */
    public static int weekNummer() {
        Calendar nu=Calendar.getInstance();
        nu.setFirstDayOfWeek(Calendar.MONDAY);
        return nu.get(Calendar.WEEK_OF_YEAR);
    }

    /**
     * geef het weeknummer van een datum
     * @return het weeknummer
     */
    public static int weekNummer(String datum) {
        Calendar cal = maakDatum(datum);
        cal.setFirstDayOfWeek(Calendar.MONDAY);
        return cal.get(Calendar.WEEK_OF_YEAR);
    }

    /**
     * Geef het nummer van de vorige week (nodig voor oude weeknr bij aanmaken nieuwe week)
     * @return
     */
    public static int vorigeWeekNummer() {
        Calendar nu=Calendar.getInstance();

        int weeknr = weekNummer();
        if (weeknr == 1) {
            return weekNummer(String.format("31-12-%04d", nu.get(Calendar.YEAR - 1)));
        } else {
            return weeknr - 1;
        }

    }

    /**
     * Geef het nummer van de vorige week (nodig voor oude weeknr bij aanmaken nieuwe week)
     * @return
     */
    public static int vorigeWeekNummer(String datum) {
        Calendar cal = maakDatum(datum);

        int weeknr = weekNummer(datum);
        if (weeknr == 1) {
            return weekNummer(String.format("31-12-%04d", cal.get(Calendar.YEAR) - 1));
        } else {
            return weeknr - 1;
        }

    }

    /**
     * geef het weekdagnummer van vandaag
     * (zondag = 1, zaterdag = 7)
     * @return
     */
    public static int weekdagNummer() {
        Calendar nu=Calendar.getInstance();
        nu.setFirstDayOfWeek(Calendar.MONDAY);
        return nu.get(Calendar.DAY_OF_WEEK);
    }

    /**
     * geef het weekdagNummer van een datum
     * @param datum (dd-mm-jjjj)
     * @return
     */
    public static int weekdagNummer(String datum) {
        Calendar cal = maakDatum(datum);
        cal.setFirstDayOfWeek(Calendar.MONDAY);
        return cal.get(Calendar.DAY_OF_WEEK);

    }

    /**
     * Geeft de datum aan de hand van weeknummer, weekdag en jaar
     * @param weeknr
     * @param weekdag vb. Calendar.FRIDAY (1=zondag, 7=zaterdag)
     * @param jaar
     * @return datum als dd-mm-jjjj
     */
    public static String datumUitWeekDag(int weeknr, int weekdag, int jaar) {
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
        Calendar cal = Calendar.getInstance();
        cal.set(Calendar.YEAR, jaar);
        cal.set(Calendar.WEEK_OF_YEAR, weeknr);
        cal.set(Calendar.DAY_OF_WEEK, weekdag);
        cal.setFirstDayOfWeek(Calendar.MONDAY);
        return sdf.format(cal.getTime());
    }

    /**
     * Geef begin [0] en einddatum [1] van de opgegeven week in het opgegeven jaar
     * (Maandag is de eerste dag van de week)
     * @param weeknr
     * @param jaar
     * @return twee datums als dd-mm-jjjj
     */
    public static String[] weekDatums(int weeknr, int jaar) {
        String[] dagen = new String[2];
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
        Calendar cal = Calendar.getInstance();
        cal.clear();
        cal.set(Calendar.WEEK_OF_YEAR, weeknr);
        cal.set(Calendar.YEAR, jaar);
        cal.setFirstDayOfWeek(Calendar.MONDAY);
        dagen[0] = sdf.format(cal.getTime());
        cal.set(Calendar.DATE, cal.get(Calendar.DATE) + 6 );
        dagen[1] = sdf.format(cal.getTime());
        return dagen;

    }

    public static String vandaag() {
        return vandaag("dd-MM-yyyy");
    }

    public static String vandaag(String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        Calendar cal = Calendar.getInstance();
        return sdf.format(cal.getTime());
    }

    /**
     * Splits een pad in een directory [0] en een filenaam [1]
     * @param pad
     * @return
     */
    public static String[] splitsPad(String pad) {
        String[] result = new String[2];
        File file = new File(pad);
        result[0] = file.getParent();
        result[1] = file.getName();
        return result;
    }


    /**
     * Geeft date object voor datum in tekst
     * @param datum
     * @param format
     * @return
     */
    public static Date stringToDate(String datum, String format) {
        Date result = null;
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            String dateInString = datum;
            result = sdf.parse(dateInString);
        } catch (Exception ex) {
        }
        return result;
    }

    /**
     * Geeft date object voor datum in tekst (jjjj-mm-dd)
     * @param datum
     * @return
     */
    public static Date stringToDate(String datum) {
        return stringToDate(datum, "dd-MM-yyyy");
    }


    /**
     * Geef de usernaam
     * @return
     */
    public static String usernaam() {
        return System.getProperty("user.name");
    }

    public static String pwd() { return System.getProperty("user.dir"); }

    /**
     * Maak een Calendar object van een datum (dd-mm-jjjj)
     * @param datum
     * @return
     */
    private static Calendar maakDatum(String datum) {
        int dag = 0, maand = 0, jaar = 0;
        Matcher mat = Pattern.compile("(\\d{2})-(\\d{2})-(\\d{4})").matcher(datum);
        while (mat.find()) {
            dag = Integer.valueOf(mat.group(1));
            maand = Integer.valueOf(mat.group(2));
            jaar = Integer.valueOf(mat.group(3));

        }
        Calendar cal=Calendar.getInstance();
        cal.set(Calendar.YEAR, jaar);
        cal.set(Calendar.MONTH, maand - 1);
        cal.set(Calendar.DAY_OF_MONTH, dag);
        cal.setFirstDayOfWeek(Calendar.MONDAY);
        return cal;
    }

}
