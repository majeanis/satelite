package cl.majeanis.satelite;

import javax.ws.rs.Consumes;
import javax.ws.rs.HeaderParam;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.context.ApplicationContext;

import cl.majeanis.satelite.util.ws.RecursoRestBase;

@Path("/sesiones")
public class Sesiones extends RecursoRestBase
{
    private static final Logger logger = LogManager.getLogger(Sesiones.class);

    @Override
    protected void initBeans(ApplicationContext appContext)
    {
    }
    

    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @POST
    public Response autenticar(@HeaderParam("Authorization") String authorization)
    {
        return null;
    }

}
