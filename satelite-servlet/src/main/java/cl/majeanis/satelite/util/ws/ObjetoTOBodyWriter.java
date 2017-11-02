package cl.majeanis.satelite.util.ws;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Type;

import javax.ws.rs.Produces;
import javax.ws.rs.WebApplicationException;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.MultivaluedMap;
import javax.ws.rs.ext.MessageBodyWriter;
import javax.ws.rs.ext.Provider;

import cl.majeanis.satelite.to.ObjetoTO;
import cl.majeanis.satelite.util.Utils;

@Provider
@Produces(MediaType.APPLICATION_JSON)
public class ObjetoTOBodyWriter implements MessageBodyWriter<ObjetoTO>
{
    @Override
    public long getSize(ObjetoTO t, Class<?> type, Type genericType, Annotation[] annotations, MediaType mediaType)
    {
        return 0;
    }

    @Override
    public boolean isWriteable(Class<?> type, Type genericType, Annotation[] annotations, MediaType mediaType)
    {
        return ObjetoTO.class.isAssignableFrom(type);
    }

    @Override
    public void writeTo(ObjetoTO t, 
            Class<?> type,
            Type genericType, 
            Annotation[] annotations, 
            MediaType mediaType,
            MultivaluedMap<String, Object> httpHeaders, 
            OutputStream entityStream) throws IOException, WebApplicationException
    {
        Utils.toJsonString(t, entityStream);
    }
}
