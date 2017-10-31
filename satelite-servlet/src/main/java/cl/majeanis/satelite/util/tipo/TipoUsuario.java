package cl.majeanis.satelite.util.tipo;

import cl.majeanis.satelite.to.modelo.TipoUsuarioTO;

public enum TipoUsuario
{
    ADMIN (1),
    CONSULTA (2)
    ;
    
    private int id;
    
    TipoUsuario(int id)
    {
        this.id = id;
    }
    
    public int getId()
    {
        return this.id;
    }
    
    public TipoUsuario of(Integer id)
    {
        for(TipoUsuario t: TipoUsuario.values())
        {
            if( t.id == id )
            {
               return t; 
            }
        }
        
        return null;
    }
    
    public TipoUsuario of(TipoUsuarioTO tipo)
    {
        if( tipo == null ) return null;
        
        return of(tipo.getId() );
    }
}
