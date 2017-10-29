/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package satelite01;

import java.io.File;
import javax.swing.filechooser.FileFilter;

/**
 *
 * @author mzavaleta
 */
public class MyFilter extends FileFilter {

    public final static int cTIPO_XLS = 0;
    public final static int cTIPO_CSV = 1;
    public final static int cTIPO_TSV = 2;
    private int tipo;

    public MyFilter(int tipo) {
        this.tipo = tipo;
    }

    public boolean accept(File file) {
        String filename = file.getName();
        if (file.isDirectory())
            return true;
        switch (tipo) {
            case cTIPO_XLS:
                return filename.toUpperCase().endsWith(".XLS");
            case cTIPO_CSV:
                return filename.toUpperCase().endsWith(".CSV");
            case cTIPO_TSV:
                return filename.toUpperCase().endsWith(".TSV");
            default:
                return filename.toUpperCase().endsWith(".CSV");
        }


    }

    public String getDescription() {
        String descrip = "";
        switch (tipo) {
            case cTIPO_XLS:
                descrip = "Hoja de cálculo Xls";
                break;
            case cTIPO_CSV:
                descrip = "Archivo separado por comas";
                break;
            case cTIPO_TSV:
                descrip = "Archivo separado por tabuladores";
                break;
            default:
                descrip = "Archivo separado por comas";
                break;
        }
        return descrip;
    }
}
