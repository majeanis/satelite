/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package Ex;

/**
 *
 * @author mzavaleta
 */
public class BU_Exception extends Exception{
    public BU_Exception(String mensaje)
    {
        super(mensaje);
    }
    public String getOraMessage() {
        String resultado = "";
        String mensFull = getMessage();
        if (mensFull.startsWith("ORA-")) {
            mensFull = mensFull.substring(11);
        }
        int pos = mensFull.indexOf("\n", 0);
        if (pos >= 0) {
            resultado = mensFull.substring(0, pos);
        } else {
            resultado = mensFull;
        }
        return resultado;
    }
}
