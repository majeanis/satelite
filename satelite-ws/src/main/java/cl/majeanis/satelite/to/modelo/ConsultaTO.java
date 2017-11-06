package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.PersistibleTO;
import cl.majeanis.satelite.util.tipo.Horario;
import cl.majeanis.satelite.util.tipo.TipoConsulta;

public class ConsultaTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;

    private String nombre;
    private String sql;
    private TipoConsulta tipo;
    private Horario horario;
    private Boolean vigente;
    private UsuarioTO dueno;
    private UsuarioTO creador;
    private ConexionTO conexion;

    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
    }
    public String getSql()
    {
        return sql;
    }
    public void setSql(String sql)
    {
        this.sql = sql;
    }
    public TipoConsulta getTipo()
    {
        return tipo;
    }
    public void setTipo(TipoConsulta tipo)
    {
        this.tipo = tipo;
    }
    public Horario getHorario()
    {
        return horario;
    }
    public void setHorario(Horario horario)
    {
        this.horario = horario;
    }
    public Boolean getVigente()
    {
        return vigente;
    }
    public void setVigente(Boolean vigente)
    {
        this.vigente = vigente;
    }
    public UsuarioTO getDueno()
    {
        return dueno;
    }
    public void setDueno(UsuarioTO dueno)
    {
        this.dueno = dueno;
    }
    public UsuarioTO getCreador()
    {
        return creador;
    }
    public void setCreador(UsuarioTO creador)
    {
        this.creador = creador;
    }
    public ConexionTO getConexion()
    {
        return conexion;
    }
    public void setConexion(ConexionTO conexion)
    {
        this.conexion = conexion;
    }
}
