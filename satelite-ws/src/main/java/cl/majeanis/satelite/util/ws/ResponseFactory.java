package cl.majeanis.satelite.util.ws;

import java.util.List;

import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.Status;

import cl.majeanis.satelite.to.ObjetoTO;
import cl.majeanis.satelite.util.JsonUtils;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;

public final class ResponseFactory
{
    public static int computeStatus(Respuesta<?> respuesta)
    {
        Resultado r = respuesta.getResult();
        if( r.isOk() )
        {
            return Status.OK.getStatusCode();
        }
        if( r.hasExceptions() )
        {
            return Status.INTERNAL_SERVER_ERROR.getStatusCode();
        }
        
        return Status.BAD_REQUEST.getStatusCode();
    }
    
    private static String toJson(Respuesta<?> respuesta, int httpStatus)
    {
        @SuppressWarnings("unused")
        class Entity
        {
            public int code;
            public String status;
            public List<String> messages;
            public ObjetoTO data;
        }
        
        Resultado r = respuesta.getResult();
        Entity e = new Entity();

        if( r.isOk() )
        {
            e.status = "success";
            e.code = (httpStatus == 0 ? Status.OK.getStatusCode(): httpStatus);
            e.messages = r.getMensajes();
        } else if ( r.hasExceptions() )
        {
            e.status = "fail";
            e.code = Status.INTERNAL_SERVER_ERROR.getStatusCode();
            e.messages = r.getErrores();
        } else
        {
            e.status = "error";
            e.code = Status.BAD_REQUEST.getStatusCode();
            e.messages = r.getErrores();
        }
        
        e.data = respuesta.getContent().orElse(null);
        return JsonUtils.toJson(e);
    }
    
    public static Response of(Respuesta<?> respuesta, int httpStatus)
    {
        if( httpStatus == 0 )
            httpStatus = computeStatus(respuesta);

        String j = toJson(respuesta, httpStatus);
        return Response.status(httpStatus).type(MediaType.APPLICATION_JSON).entity(j).build();
    }
    
    public static Response of(Respuesta<?> respuesta)
    {
        return of(respuesta,0);
    }
    
    public static Response of(Exception exception)
    {
        Resultado rtdo = new ResultadoProceso();
        rtdo.addError(exception);
        return of( new Respuesta<ObjetoTO>(rtdo));
    }
}
