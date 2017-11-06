package cl.majeanis.satelite;

import java.math.BigInteger;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;
import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.HeaderParam;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.PathParam;
import javax.ws.rs.Produces;
import javax.ws.rs.core.Context;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.context.ApplicationContext;

import cl.majeanis.satelite.po.ConexionPO;
import cl.majeanis.satelite.to.modelo.ConexionTO;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.ws.RecursoRestBase;
import cl.majeanis.satelite.util.ws.ResponseFactory;

@Path("/conexiones")
public class Conexiones extends RecursoRestBase
{
    private static final Logger logger = LogManager.getLogger(Conexiones.class);
    
    private ConexionPO conxPO;

    private static int numeroConexion = 1;
    
    @Override
    protected void initBeans(ApplicationContext appContext)
    {
        conxPO = appContext.getBean(ConexionPO.class);
    }

    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @POST
    public Response guardar(@HeaderParam("Authorization") String authorization,
                            @Context HttpServletRequest request,
                            ConexionTO data)
    {
        logger.info("guardar[INI] authorization={} data={}", authorization, data );
        conxPO.guardar(data);
        Respuesta<ConexionTO> r = new Respuesta<ConexionTO>(data);

        HttpSession s = request.getSession();
        s.setAttribute("conexion", data);
        return ResponseFactory.of(r);
    }

    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response getConexion(@HeaderParam("Authorization") String authorization, 
                                @Context HttpServletRequest request)
    {
        logger.info("getConexion[INI] authorization={}", authorization );

        HttpSession s = request.getSession();
        ConexionTO d = (ConexionTO) s.getAttribute("conexion");

        Respuesta<ConexionTO> r = new Respuesta<>(d);
        return ResponseFactory.of(r);
    }

    @Path("{idConexion}")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response get(@HeaderParam("Authorization") String authorization,
                        @PathParam("idConexion") BigInteger idConexion,
                        @Context HttpServletRequest request)
    {
        HttpSession s = request.getSession();
        ConexionTO data = (ConexionTO) s.getAttribute("data");

        return Response.ok(data).build();
    }
}
