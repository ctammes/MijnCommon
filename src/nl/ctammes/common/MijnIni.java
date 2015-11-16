package nl.ctammes.common;

import org.ini4j.Ini;
import org.ini4j.Wini;

import javax.swing.*;
import java.io.File;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 26-3-13
 * Time: 20:12
 * To change this template use File | Settings | File Templates.
 */
public class MijnIni {

    private String inifile = "";

    /**
     * Maak inifile aan indien nodig
     * @param inifile
     */
    public MijnIni(String inifile) {
        File file = new File(inifile);
        if (!file.exists()) {
            if (JOptionPane.showConfirmDialog(null, "Inifile " + inifile + " niet gevonden. Aanmaken?", "Bevestig keuze", JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION) {
                try {
                    file.createNewFile();
                } catch(Exception e) {
                    System.out.println(e.getMessage());
                }
            }
        }
        this.inifile = inifile;
    }


    /**
     * Lees waarde uit sleutel in sectie
     * @param section
     * @param key
     * @return
     */
    public String lees(String section, String key) {

        String value = "";
        try{
            File file = new File(inifile);
            if (file.exists()) {
                Ini ini = new Ini(file);
                value = ini.get(section, key);
            } else {
                throw new Exception(inifile + " niet gevonden!");
            }
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
        return value;
    }

    /**
     * Lees waarde uit sleutel in sectie
     * @param section
     * @param key
     * @param def   default waarde
     * @return
     */
    public String lees(String section, String key, String def) {

        String value = lees(section, key);
        if (value == null) {
            value = def;
        }
        return value;
    }

    /**
     * Schrijf waarde naar sleutel in sectie
     * @param section
     * @param key
     * @param value
     */
    public void schrijf(String section, String key, String value) {

        try{
            File file = new File(inifile);
            if (file.exists()) {
                Wini ini = new Wini(new File(inifile));
                ini.put(section, key, value);
                ini.store();
            }
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

}
