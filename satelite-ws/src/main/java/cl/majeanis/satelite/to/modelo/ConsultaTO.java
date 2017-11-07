package cl.majeanis.satelite.to.modelo;

import java.time.LocalDateTime;

import cl.majeanis.satelite.to.PersistibleTO;
import cl.majeanis.satelite.util.tipo.Horario;
import cl.majeanis.satelite.util.tipo.TipoConsulta;

public class ConsultaTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;

    private String nombre;
    private ConexionTO conexion;
    private String sql;    
    private UsuarioTO dueno;
    private UsuarioTO creador;
    private LocalDateTime creacion;
    private LocalDateTime ultActualizacion;
    private Horario horario;
    private TipoConsulta tipo;
    private Boolean vigente;

    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
    }
    public ConexionTO getConexion()
    {
        return conexion;
    }
    public void setConexion(ConexionTO conexion)
    {
        this.conexion = conexion;
    }
    public String getSql()
    {
        return sql;
    }
    public void setSql(String sql)
    {
        this.sql = sql;
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
    public LocalDateTime getCreacion()
    {
        return creacion;
    }
    public void setCreacion(LocalDateTime creacion)
    {
        this.creacion = creacion;
    }
    public LocalDateTime getUltActualizacion()
    {
        return ultActualizacion;
    }
    public void setUltActualizacion(LocalDateTime ultActualizacion)
    {
        this.ultActualizacion = ultActualizacion;
    }
    public Horario getHorario()
    {
        return horario;
    }
    public void setHorario(Horario horario)
    {
        this.horario = horario;
    }
    public TipoConsulta getTipo()
    {
        return tipo;
    }
    public void setTipo(TipoConsulta tipo)
    {
        this.tipo = tipo;
    }
    public Boolean getVigente()
    {
        return vigente;
    }
    public void setVigente(Boolean vigente)
    {
        this.vigente = vigente;
    }
}
