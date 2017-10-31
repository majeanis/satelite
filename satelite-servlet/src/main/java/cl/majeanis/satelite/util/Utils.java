package cl.majeanis.satelite.util;

import java.util.List;

import cl.majeanis.satelite.util.json.GsonFacade;

public class Utils
{
    public static String toJsonString(Object obj)
    {
        return GsonFacade.toJsonString(obj);
    }
    
    public static <T> T fromJsonString(String json, Class<T> clazz)
    {
        return GsonFacade.fromJsonString(json, clazz );
    }
    
    public static int sizeOf(List<?> lista)
    {
        if( lista == null )
            return 0;
        
        return lista.size();
    }
}
