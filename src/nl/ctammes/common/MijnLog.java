package nl.ctammes.common;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.*;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 17-3-13
 * Time: 11:30
 * To change this template use File | Settings | File Templates.
 */
public class MijnLog {

    private String logdir;
    private String logNaam;
    private Logger log = null;

    public String getLogdir() {
        return logdir;
    }

    public String getLogNaam() {
        return logNaam;
    }

    public MijnLog(String logdir, String logNaam, boolean append) {
        this.logdir = logdir;
        this.logNaam = logNaam;

        try {
            // console handler
            ConsoleHandler cons = new ConsoleHandler();

            // filehandler (logfile) - append to log (=true)
            FileHandler hand = new FileHandler(logdir + "/" + logNaam, append);

            // formattering logregel via aangepaste class
            MijnSimpleFormatter formatterTxt = new MijnSimpleFormatter();
            hand.setFormatter(formatterTxt);
            cons.setFormatter(formatterTxt);
            // cons.setLevel(Level.OFF);    // geen output naar console

            log = Logger.getLogger(MijnLog.class.getName());
            log.setUseParentHandlers(false);    // zet default console log uit
            log.addHandler(hand);
            log.addHandler(cons);
        } catch (Exception e) {
//            e.printStackTrace();
            throw new RuntimeException("Probleem bij het maken van de logfile: " + logdir + "/" + logNaam);
        }

    }

    public Logger getLog() {
        return log;
    }
}


/**
 * Op deze manier kun je de logregel zelf formatteren
 */
class MijnSimpleFormatter extends SimpleFormatter {

    // oorspronkelijk formaat (over 2 regels):
    // mrt 17, 2013 4:52:46 PM TestDb testLog
    // INFO: dit is een test
    // nu:
    // 2013-03-17 17:21:57 TestDb testLog - INFO: dit is een test
    public String format(LogRecord record) {
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String timestamp = df.format(new Date());
        return String.format("%s %s %s - %s: %s\n", timestamp, record.getSourceClassName(), record.getSourceMethodName(), record.getLevel(), record.getMessage());
    }
}