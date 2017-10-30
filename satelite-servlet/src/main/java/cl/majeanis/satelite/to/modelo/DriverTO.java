package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.PersistibleTO;

public class DriverTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;

    private String nombre;
    private String clazz;

    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
    }
    public String getClazz()
    {
        return clazz;
    }
    public void setClazz(String clazz)
    {
        this.clazz = clazz;
    }
}
