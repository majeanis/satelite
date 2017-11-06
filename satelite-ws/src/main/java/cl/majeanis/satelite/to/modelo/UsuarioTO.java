package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.PersistibleTO;

public class UsuarioTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;
    
    private String login;
    private String nombre;
    private Boolean vigente;
    private TipoUsuarioTO tipo;

    public String getLogin()
    {
        return login;
    }
    public void setLogin(String login)
    {
        this.login = login;
    }
    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
    }
    public Boolean getVigente()
    {
        return vigente;
    }
    public void setVigente(Boolean vigente)
    {
        this.vigente = vigente;
    }
    public TipoUsuarioTO getTipo()
    {
        return tipo;
    }
    public void setTipo(TipoUsuarioTO tipo)
    {
        this.tipo = tipo;
    }
}

