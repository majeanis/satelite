package cl.majeanis.satelite.util;

import java.util.List;

public class Utils
{
    public static int sizeOf(List<?> lista)
    {
        if( lista == null )
            return 0;
        
        return lista.size();
    }
}
