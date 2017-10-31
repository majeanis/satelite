package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.BaseTO;

public class TipoUsuarioTO extends BaseTO
{
    private static final long serialVersionUID = 1L;
    
    private Integer id;
    private String nombre;
    private Boolean administrador;
    private Boolean creaConsulta;
    private Boolean autoAsignarConsulta;
    private Boolean modificrConsulta;
    private Boolean eliminarConsulta;
    private Boolean ejecutarConsulya;

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
    public Boolean getAdministrador()
    {
        return administrador;
    }
    public void setAdministrador(Boolean administrador)
    {
        this.administrador = administrador;
    }
    public Boolean getCreaConsulta()
    {
        return creaConsulta;
    }
    public void setCreaConsulta(Boolean creaConsulta)
    {
        this.creaConsulta = creaConsulta;
    }
    public Boolean getAutoAsignarConsulta()
    {
        return autoAsignarConsulta;
    }
    public void setAutoAsignarConsulta(Boolean autoAsignarConsulta)
    {
        this.autoAsignarConsulta = autoAsignarConsulta;
    }
    public Boolean getModificrConsulta()
    {
        return modificrConsulta;
    }
    public void setModificrConsulta(Boolean modificrConsulta)
    {
        this.modificrConsulta = modificrConsulta;
    }
    public Boolean getEliminarConsulta()
    {
        return eliminarConsulta;
    }
    public void setEliminarConsulta(Boolean eliminarConsulta)
    {
        this.eliminarConsulta = eliminarConsulta;
    }
    public Boolean getEjecutarConsulya()
    {
        return ejecutarConsulya;
    }
    public void setEjecutarConsulya(Boolean ejecutarConsulya)
    {
        this.ejecutarConsulya = ejecutarConsulya;
    }
}