package cl.majeanis.satelite.util;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;
import java.util.List;

import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;

public class ToStringUtils
{
    public static String toString(Object objeto)
    {
        return ToStringBuilder.reflectionToString(objeto, ToStringStyle.SHORT_PREFIX_STYLE, true);
    }

    public static String toString(Exception exception)
    {
        if (exception == null)
            return "";

        Writer writer = new StringWriter();
        PrintWriter printWriter = new PrintWriter(writer);

        exception.printStackTrace(printWriter);
        return writer.toString();
    }

    public static String toString(Long value)
    {
        if (value == null)
            return "";

        return String.valueOf(value);
    }

    public static String toString(String cadena)
    {
        if (cadena == null)
            return "";
        return cadena;
    }
    
    public static String toString(List<String> lista)
    {
        if( Utils.sizeOf(lista) == 0 )
            return "";

        StringBuilder sb = new StringBuilder();
        for(String obj: lista)
        {
            sb.append( "," + obj);
        }

        return "[" + sb.substring(1) + "]";
    }
}
