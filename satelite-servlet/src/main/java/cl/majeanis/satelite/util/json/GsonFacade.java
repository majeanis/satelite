package cl.majeanis.satelite.util.json;

import java.time.LocalTime;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

public class GsonFacade
{
    private static Gson stringWriter;
    
    private static Gson getStringWriter()
    {
        if( stringWriter != null )
            return stringWriter;
        
        synchronized( GsonFacade.class )
        {
            if( stringWriter != null )
                return stringWriter;
            
            GsonBuilder bldr = new GsonBuilder();
            bldr.registerTypeAdapter(LocalTime.class, new LocalTimeConverter());
            
            stringWriter = bldr.create();
            return stringWriter;
        }
    }
    
    public static String toJsonString(Object obj)
    {
        return getStringWriter().toJson(obj);
    }
    
    public static <T> T fromJsonString(String json, Class<T> clazz)
    {
        return getStringWriter().fromJson(json, clazz);
    }
}
