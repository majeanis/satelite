package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.BaseTO;

public class DriverTO extends BaseTO
{
    private static final long serialVersionUID = 1L;

    private Integer id;
    private String nombre;
    private String clazz;

    public Integer getId()
    {
        return id;
    }
    public void setId(Integer id)
    {
        this.id = id;
    }
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
