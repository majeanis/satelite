package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.PersistibleTO;
import cl.majeanis.satelite.util.tipo.TipoAyuda;

public class ParametroTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;

    private String nombre;
    private String texto;
    private String type;
    private TipoAyuda tipoAyuda;
    private String valores;
    private Boolean opcional;
    private ConsultaTO consulta;

    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
    }
    public String getTexto()
    {
        return texto;
    }
    public void setTexto(String texto)
    {
        this.texto = texto;
    }
    public String getType()
    {
        return type;
    }
    public void setType(String type)
    {
        this.type = type;
    }
    public TipoAyuda getTipoAyuda()
    {
        return tipoAyuda;
    }
    public void setTipoAyuda(TipoAyuda tipoAyuda)
    {
        this.tipoAyuda = tipoAyuda;
    }
    public String getValores()
    {
        return valores;
    }
    public void setValores(String valores)
    {
        this.valores = valores;
    }
    public Boolean getOpcional()
    {
        return opcional;
    }
    public void setOpcional(Boolean opcional)
    {
        this.opcional = opcional;
    }
    public ConsultaTO getConsulta()
    {
        return consulta;
    }
    public void setConsulta(ConsultaTO consulta)
    {
        this.consulta = consulta;
    }
}
