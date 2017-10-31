package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.PersistibleTO;
import cl.majeanis.satelite.util.tipo.Encrypted;

public class ConexionTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;

    private String nombre;
    private String url;
    private String usuario;
    private Encrypted password;
    private DriverTO driver;

    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
    }
    public String getUrl()
    {
        return url;
    }
    public void setUrl(String url)
    {
        this.url = url;
    }
    public String getUsuario()
    {
        return usuario;
    }
    public void setUsuario(String usuario)
    {
        this.usuario = usuario;
    }
    public Encrypted getPassword()
    {
        return password;
    }
    public void setPassword(Encrypted password)
    {
        this.password = password;
    }
    public DriverTO getDriver()
    {
        return driver;
    }
    public void setDriver(DriverTO driver)
    {
        this.driver = driver;
    }
}
