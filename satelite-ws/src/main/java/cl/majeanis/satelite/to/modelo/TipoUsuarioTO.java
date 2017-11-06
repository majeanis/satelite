package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.BaseTO;

public class TipoUsuarioTO extends BaseTO
{
    private static final long serialVersionUID = 1L;
    
    private String codigo;
    private Boolean administrador = false;
    private Boolean creaConsulta = false;
    private Boolean autoAsignarConsulta = false;
    private Boolean modificarConsulta = false;
    private Boolean eliminarConsulta = false;
    private Boolean ejecutarConsulta = false;

    public String getCodigo()
    {
        return codigo;
    }
    public void setCodigo(String codigo)
    {
        this.codigo = codigo;
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
    public Boolean getModificarConsulta()
    {
        return modificarConsulta;
    }
    public void setModificarConsulta(Boolean modificarConsulta)
    {
        this.modificarConsulta = modificarConsulta;
    }
    public Boolean getEliminarConsulta()
    {
        return eliminarConsulta;
    }
    public void setEliminarConsulta(Boolean eliminarConsulta)
    {
        this.eliminarConsulta = eliminarConsulta;
    }
    public Boolean getEjecutarConsulta()
    {
        return ejecutarConsulta;
    }
    public void setEjecutarConsulta(Boolean ejecutarConsulya)
    {
        this.ejecutarConsulta = ejecutarConsulya;
    }
}
