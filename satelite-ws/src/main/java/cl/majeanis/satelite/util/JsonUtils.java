package cl.majeanis.satelite.util;

import java.io.InputStream;
import java.io.OutputStream;

import cl.majeanis.satelite.util.json.JacksonFacade;

public final class JsonUtils
{
    public static String toJson(Object object)
    {
        return JacksonFacade.toJson(object);
    }

    public static void toJson(Object object, OutputStream stream)
    {
        JacksonFacade.toJson(object, stream);
    }

    public static <T> T fromJson(Class<T> clazz, String json)
    {
        return JacksonFacade.fromJson(clazz, json);
    }

    public static <T> T fromJson(Class<T> clazz, InputStream stream)
    {
        return JacksonFacade.fromJson(clazz, stream);
    }
}
