package cl.majeanis.satelite;

import java.util.List;

import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.HeaderParam;
import javax.ws.rs.Path;
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
    
    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @GET
    public Response get(@HeaderParam("Authorization") String authorization)
    {
        logger.info("get[INI] authorization={}", authorization);
        
        Respuesta<List<ConsultaTO>> r = consBO.getList(sesion);
        
        logger.info("get[FIN] respuesta={}", r);
        return ResponseFactory.of(r);
    }

//    @Path("{usuario}")
//    @Produces(MediaType.APPLICATION_JSON)
//    @Consumes(MediaType.APPLICATION_JSON)
//    @GET
//    public Response get(@HeaderParam("Authorization") String authorization,
//                        @PathParam("usuario") String usuario)
//    {
//        logger.info("get[INI] authorization={} usuario={}", authorization, usuario);
//        
//        Respuesta<List<ConsultaTO>> r = consBO.getList(sesion, usuario);
//        
//        logger.info("get[FIN] authorization={} usuario");
//        return ResponseFactory.of(r);
//    }
}
