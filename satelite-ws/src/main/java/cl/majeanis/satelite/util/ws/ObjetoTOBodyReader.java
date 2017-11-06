package cl.majeanis.satelite.util.ws;

import java.io.IOException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Type;

import javax.ws.rs.Consumes;
import javax.ws.rs.WebApplicationException;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.MultivaluedMap;
import javax.ws.rs.core.NoContentException;
import javax.ws.rs.ext.MessageBodyReader;
import javax.ws.rs.ext.Provider;

import cl.majeanis.satelite.to.ObjetoTO;
import cl.majeanis.satelite.util.JsonUtils;

@Provider
@Consumes(MediaType.APPLICATION_JSON)
public class ObjetoTOBodyReader implements MessageBodyReader<ObjetoTO>
{
    @Override
    public boolean isReadable(Class<?> type, Type genericType, Annotation[] annotations, MediaType mediaType)
    {
        return ObjetoTO.class.isAssignableFrom(type);
    }

    @Override
    public ObjetoTO readFrom(Class<ObjetoTO> type, 
            Type genericType, 
            Annotation[] annotations, 
            MediaType mediaType,
            MultivaluedMap<String, String> httpHeaders, 
            InputStream entityStream)  throws IOException, WebApplicationException
    {
        //
        // Eventualmente el InputStream contiene acceso a un objeto vacío, en cuyo caso
        // se producirán Exceptions durante la deserialización. Para evitar lo anterior
        // verificaremos que el InputStream contenga información,  mediante la revisión
        // del HEADER "Content-Length", el cual debe tener un valor mayor que cero.
        String contentLength = httpHeaders.getFirst("content-length");
        int bytesEstimaded = entityStream.available();

        if( "0".equals(contentLength) && bytesEstimaded == 0)
            throw new NoContentException("Objeto vacío");

        ObjetoTO o = JsonUtils.fromJson(type, entityStream );
        if( o == null )
        {
            throw new NoContentException("Objeto no pudo ser deserializado");
        }

        return o;
    }
}
