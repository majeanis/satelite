/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package logger;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import satelite01.Constantes;

/**
 *
 * @author mzavaleta
 */
public class Log {

    private static String log_file;
    private static String log_level;

    static {
        log_file = Constantes.propiedades.getProperty("log_file", "jsatelite.log");
        log_level = Constantes.propiedades.getProperty("log_level", "info");
    }

    private static void escribir(String mensaje, FileOutputStream fos, String level) {
        try {
            fos.write(setMsje(mensaje, level));
        } catch (FileNotFoundException ex1) {
        } catch (IOException ex1) {
        }
    }

    private static byte[] setMsje(String mensaje, String level) {
        Date fecha = new Date();
        StringBuilder linea = new StringBuilder();
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        linea.append(sdf.format(fecha));
        linea.append(" : ");
        linea.append(level);
        linea.append(" : ");
        linea.append(mensaje);
        linea.append("\r\n");
        return linea.toString().getBytes();
    }

    private static void escribir(String mensaje, String level) {
        try {
            FileOutputStream fos = new FileOutputStream(new File(log_file), true);
            fos.write(setMsje(mensaje, level));
            fos.close();
        } catch (FileNotFoundException ex1) {
        } catch (IOException ex1) {
        }
    }

    public static void error(String mensaje, Error err) {
        if (log_level.equalsIgnoreCase("debug") || log_level.equalsIgnoreCase("info") || log_level.equalsIgnoreCase("error")) {
            File log = new File(log_file);
            try {
                FileOutputStream fos = new FileOutputStream(log, true);
                escribir(mensaje, fos,"error");
                if (err != null) {
                    PrintStream ps = new PrintStream(fos);
                    err.printStackTrace(ps);
                    ps.close();
                }
                fos.close();
            } catch (FileNotFoundException ex1) {
            } catch (IOException ex1) {
            }
        }
    }

    public static void debug(String mensaje) {
        if (log_level.equalsIgnoreCase("debug")) {
            escribir(mensaje, "debug");
        }
    }

    public static void info(String mensaje) {
        if (log_level.equalsIgnoreCase("debug") || log_level.equalsIgnoreCase("info")) {
            escribir(mensaje, "info");
        }
    }

    public static void error(String mensaje, Exception ex) {
        if (log_level.equalsIgnoreCase("debug") || log_level.equalsIgnoreCase("info") || log_level.equalsIgnoreCase("error")) {

            File log = new File(log_file);
            try {
                FileOutputStream fos = new FileOutputStream(log, true);
                escribir(mensaje, fos, "error");
                if (ex != null) {
                    PrintStream ps = new PrintStream(fos);
                    ex.printStackTrace(ps);
                    ps.close();
                }
                fos.close();
            } catch (FileNotFoundException ex1) {
            } catch (IOException ex1) {
            }
        }
    }
}
