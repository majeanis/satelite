package cl.majeanis.satelite.to.modelo;

import java.time.LocalDateTime;

import cl.majeanis.satelite.to.PersistibleTO;

public class LogTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;

    private UsuarioTO usuario;
    private ConsultaTO consulta;
    private LocalDateTime fechaInicio;
    private LocalDateTime fechaTermino;
    private int tiempo;
    private int registros;

    public UsuarioTO getUsuario()
    {
        return usuario;
    }
    public void setUsuario(UsuarioTO usuario)
    {
        this.usuario = usuario;
    }
    public ConsultaTO getConsulta()
    {
        return consulta;
    }
    public void setConsulta(ConsultaTO consulta)
    {
        this.consulta = consulta;
    }
    public LocalDateTime getFechaInicio()
    {
        return fechaInicio;
    }
    public void setFechaInicio(LocalDateTime fechaInicio)
    {
        this.fechaInicio = fechaInicio;
    }
    public LocalDateTime getFechaTermino()
    {
        return fechaTermino;
    }
    public void setFechaTermino(LocalDateTime fechaTermino)
    {
        this.fechaTermino = fechaTermino;
    }
    public int getTiempo()
    {
        return tiempo;
    }
    public void setTiempo(int tiempo)
    {
        this.tiempo = tiempo;
    }
    public int getRegistros()
    {
        return registros;
    }
    public void setRegistros(int registros)
    {
        this.registros = registros;
    }
}
