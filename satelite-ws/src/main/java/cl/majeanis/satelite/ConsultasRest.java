package cl.majeanis.satelite;

import java.math.BigInteger;
import java.util.List;
import java.util.Optional;

import javax.ws.rs.GET;
import javax.ws.rs.HeaderParam;
import javax.ws.rs.Path;
import javax.ws.rs.PathParam;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;

import cl.majeanis.satelite.bo.ConsultaBO;
import cl.majeanis.satelite.to.modelo.ConsultaTO;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.ws.RecursoRestBase;
import cl.majeanis.satelite.util.ws.ResponseFactory;

@Path("/consultas")
public class ConsultasRest extends RecursoRestBase
{
    private static final Logger logger = LogManager.getLogger(ConsultasRest.class);

    @Autowired
    private ConsultaBO consBO;
    
    @Override
    protected void initBeans(ApplicationContext appContext)
    {
        consBO = appContext.getBean(ConsultaBO.class);
    }
    
    @Path("usuarios/{usuario}")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response consultas(@HeaderParam("Authorization") String authorization,
                              @PathParam("usuario") String usuario)
    {
        logger.info("consultas[INI] authorization={} usuario={}", authorization, usuario);
        
        Respuesta<List<ConsultaTO>> r = consBO.getList(sesionRequest, Optional.ofNullable(usuario) );
        
        logger.info("consultas[FIN] authorization={} respuesta={}", r);
        return ResponseFactory.of(r);
    }
    
    @Path("{consultaId}")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response consulta(@HeaderParam("Authorization") String authorization,
                             @PathParam("consultaId") BigInteger consultaId)
    {
        logger.info("consulta[INI] authorization={} consultaId={}", authorization, consultaId);
        
        Respuesta<ConsultaTO> r = consBO.get(sesionRequest, consultaId);
        
        logger.info("consulta[FIN] consultaId={} respuesta={}", consultaId, r);
        return ResponseFactory.of(r);
    }
    
    @Path("{consultaId}/ejecutar")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response ejecutar(@HeaderParam("Authorization") String authorization,
                             @PathParam("consultaId") BigInteger consultaId)
    {
        logger.info("ejecutar[INI] authorization={} consultaId={}", authorization, consultaId );
        
        Respuesta<ConsultaTO> r = consBO.ejecutar(sesionRequest, consultaId);
        
        logger.info("ejecutar[FIN] consultaId={} r={}", consultaId, r);
        return ResponseFactory.of(r);
    }
    
}
