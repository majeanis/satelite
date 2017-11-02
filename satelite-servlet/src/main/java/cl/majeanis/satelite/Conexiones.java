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
import javax.ws.rs.core.Response.Status;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.context.ApplicationContext;

import cl.majeanis.satelite.po.ConexionPO;
import cl.majeanis.satelite.to.modelo.ConexionTO;
import cl.majeanis.satelite.util.tipo.Encrypted;
import cl.majeanis.satelite.util.ws.RecursoRestBase;

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
    public Response guardar(@HeaderParam("X-Sesion") String sesion)
    {
        ConexionTO data = new ConexionTO();
        logger.info("guardar[INI] sesion={} data={}", sesion, data );

        conxPO.guardar(data);
        return Response.status(Status.OK).entity(data).build();
    }

    @Path("")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response getConexion(@HeaderParam("X-Sesion") String sesion, @Context HttpServletRequest request)
    {
        logger.info("getConexion[INI] sesion={}", sesion );

        ConexionTO data = new ConexionTO();
        data.setNombre("Conexion " + numeroConexion++ );
        data.setUsuario( new Encrypted( "el usuario"));
        data.setUrl( new Encrypted("la URL"));
        data.setPassword( new Encrypted("la pass"));

        conxPO.guardar(data);
        
        HttpSession s = request.getSession();
        s.setAttribute("data", data);

        return Response.ok(data).build();
    }

    @Path("{idConexion}")
    @Produces(MediaType.APPLICATION_JSON)
    @GET
    public Response get(@HeaderParam("X-Sesion") String sesion,
                        @PathParam("idConexion") BigInteger idConexion,
                        @Context HttpServletRequest request)
    {
        HttpSession s = request.getSession();
        ConexionTO data = (ConexionTO) s.getAttribute("data");

        return Response.ok(data).build();
    }
}
