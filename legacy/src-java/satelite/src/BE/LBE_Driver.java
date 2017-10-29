/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BE;

/**
 *
 * @author mzavaleta
 */
public class LBE_Driver {
    private int codigo;
    private String nombre;
    private String driver;
    private String url;

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public int getCodigo() {
        return codigo;
    }

    public void setCodigo(int codigo) {
        this.codigo = codigo;
    }

    public String getDriver() {
        return driver;
    }

    public void setDriver(String driver) {
        this.driver = driver;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
    @Override
    public String toString(){
        return this.nombre + " (" + this.driver + ")";
    }
}
