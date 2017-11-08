package cl.majeanis.satelite;

import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.HeaderParam;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.context.ApplicationContext;

import cl.majeanis.satelite.bo.SesionBO;
import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.ws.RecursoRestBase;
import cl.majeanis.satelite.util.ws.ResponseFactory;

@Path("/sesiones")
public class SesionesRest extends RecursoRestBase
{
    private static final Logger logger = LogManager.getLogger(SesionesRest.class);

    private SesionBO sesionBO;
    
    @Override
    protected void initBeans(ApplicationContext appContext)
    {
        sesionBO = appContext.getBean(SesionBO.class);
    }

    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @POST
    public Response autenticar(@HeaderParam("Authorization") String authorization)
    {
        logger.info("autenticar[INI] authorization={}", authorization );

        try
        {
            Respuesta<SesionTO> r = sesionBO.autenticar(authorization);
            logger.info("autenticar[FIN] respuesta={}", r );
            
            return ResponseFactory.of(r);
        } catch(Exception e)
        {
            logger.error("autenticar[ERR]", e);
            return ResponseFactory.of(e);
        }
    }

    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @GET
    public Response checkToken(@HeaderParam("Authorization") String authorization)
    {
        logger.info("checkToken[INI] authorization={}", authorization );

        try
        {
            Respuesta<SesionTO> r = sesionBO.obtener(authorization);
            logger.info("checkToken[FIN] respuesta={}", r );
            
            return ResponseFactory.of(r);
        } catch(Exception e)
        {
            logger.error("checkToken[ERR]", e);
            return ResponseFactory.of(e);
        }
    }
}
