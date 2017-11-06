package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.PersistibleTO;

public class UsuarioTO extends PersistibleTO
{
    private static final long serialVersionUID = 1L;
    
    private String nombre;
    private TipoUsuarioTO tipo;

    public String getNombre()
    {
        return nombre;
    }
    public void setNombre(String nombre)
    {
        this.nombre = nombre;
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

