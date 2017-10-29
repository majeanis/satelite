/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BE;

/**
 *
 * @author mzavaleta
 */
public class LBE_BDatos {
    protected LBE_Driver Driver;

    public LBE_Driver getDriver() {
        return Driver;
    }

    public void setDriver(LBE_Driver Driver) {
        this.Driver = Driver;
    }
    protected String nombre;

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    protected int codigo;

    public int getCodigo() {
        return codigo;
    }

    public void setCodigo(int codigo) {
        this.codigo = codigo;
    }

    public String getClave() {
        return clave;
    }

    public void setClave(String clave) {
        this.clave = clave;
    }


    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public String getUsuario() {
        return usuario;
    }

    public void setUsuario(String usuario) {
        this.usuario = usuario;
    }
    protected String url;
    protected String usuario;
    protected String clave;
    @Override
    public String toString()
    {
        return this.nombre + "(" + this.usuario + "@" + this.url+ ")" ;
    } 
}
