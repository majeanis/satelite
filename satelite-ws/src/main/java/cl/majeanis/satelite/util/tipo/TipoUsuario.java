package cl.majeanis.satelite.util.tipo;

import cl.majeanis.satelite.to.modelo.TipoUsuarioTO;

public enum TipoUsuario
{
    ADMIN,
    CONSULTA
    ;
    
    public static TipoUsuario of(TipoUsuarioTO tipo)
    {
        if( tipo == null ) return null;
        
        return valueOf(tipo.getNombre());
    }
}
