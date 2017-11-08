package cl.majeanis.satelite.to.modelo;

import java.time.LocalDateTime;

import com.fasterxml.jackson.annotation.JsonIgnore;

import cl.majeanis.satelite.to.BaseTO;

public class SesionTO extends BaseTO
{
    private static final long serialVersionUID = 1L;

    private String id;
    private UsuarioTO usuario;
    private LocalDateTime fecha;

    public String getId()
    {
        return id;
    }
    public void setId(String id)
    {
        this.id = id;
    }
    public UsuarioTO getUsuario()
    {
        return usuario;
    }
    public void setUsuario(UsuarioTO usuario)
    {
        this.usuario = usuario;
    }
    public LocalDateTime getFecha()
    {
        return fecha;
    }
    public void setFecha(LocalDateTime fecha)
    {
        this.fecha = fecha;
    }
    
    @JsonIgnore
    public boolean isAdmin()
    {
        if( usuario == null )
            return false;
        
        TipoUsuarioTO tipo = usuario.getTipo();
        if( tipo == null )
            return false;
        
        if( tipo.getAdministrador() == false )
            return false;
        
        return tipo.getAdministrador();
    }
    
    @JsonIgnore
    public String getNombreUsuario()
    {
        if( usuario == null )
            return "";
        return usuario.getNombre();
    }
}
