package nl.ctammes.common;/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

import java.sql.*;
import java.util.logging.Logger;

/**
 *
 * @author chris
 */
public class Sqlite {

    private String sqliteDir = null;
    private String sqliteDb = null;
    private Connection conn = null;
    private int queryTimeout = 30;
    private Logger log = null;

    public void setQueryTimeout(int queryTimeout) {
        this.queryTimeout = queryTimeout;
    }

    public Sqlite(String dir, String db) {
        this.sqliteDir = dir;
        this.sqliteDb = db;
    }

    public Sqlite(String dir, String db, Logger log) {
        this.sqliteDir = dir;
        this.sqliteDb = db;
        this.log = log;
    }

    /**
     * SQLite maakt niet bestaande database aan
     */
    public void openDb() {
        try {
            Class.forName("org.sqlite.JDBC");

            // create a database connection
            String cn = "jdbc:sqlite:" + this.sqliteDir + "/" + this.sqliteDb;
            this.conn = DriverManager.getConnection(cn);
        } catch(SQLException e) {
            String msg = e.getClass().toString() + " : "+e.getMessage();
//            log.severe(msg);
            throw new RuntimeException(msg);
        } catch(ClassNotFoundException e) {
            String msg = e.getClass().toString() + " : "+e.getMessage();
//            log.severe(msg);
            throw new RuntimeException(msg);
        }
    }
    
    public void sluitDb() {
        try {
            if(this.conn != null) {
              this.conn.close();
            }
        } catch(SQLException e) {
            log.severe(e.getMessage());
        }
    }

    /**
     * Execute statement
     * @param sql
     * @return resultset
     */
    public ResultSet execute(String sql) {
        try {
            Statement statement = this.conn.createStatement();
            statement.setQueryTimeout(queryTimeout);  // set timeout to 30 sec.

            ResultSet rs = statement.executeQuery(sql);
            return rs;
        } catch(OutOfMemoryError e) {
            // if the error message is "out of memory",
            // it probably means no database file is found
            String msg = "Bestaat de database wel?";
//            log.severe(msg);
            throw new RuntimeException(msg);
        } catch(SQLException e) {
            String msg = e.getClass().toString() + " : "+e.getMessage();
//            log.severe(msg);
            throw new RuntimeException(msg);
        }
        
    }


    /**
     * Execute statement zonder resultset (create/drop/insert/update)
     * @param sql
     * @return true/false
     */
    public boolean executeNoResult(String sql) {
        try {
            Statement statement = this.conn.createStatement();
            statement.setQueryTimeout(queryTimeout);  // set timeout to 30 sec.
            statement.executeUpdate(sql);
            return true;
        } catch(OutOfMemoryError e) {
            // if the error message is "out of memory",
            // it probably means no database file is found
            String msg = "Bestaat de database wel?";
//            log.severe(msg);
            throw new RuntimeException(msg);
        } catch(SQLException e) {
            String msg = e.getClass().toString() + " : "+e.getMessage();
//            log.severe(msg);
            throw new RuntimeException(msg);
        }
        
    }


}
