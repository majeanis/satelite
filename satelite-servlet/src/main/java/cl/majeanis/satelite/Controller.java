package cl.majeanis.satelite;

import java.io.IOException;

import javax.servlet.ServletConfig;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.context.ApplicationContext;
import org.springframework.web.context.support.WebApplicationContextUtils;

import cl.majeanis.satelite.po.ConexionPO;
import cl.majeanis.satelite.to.modelo.ConexionTO;
import cl.majeanis.satelite.to.modelo.DriverTO;
import cl.majeanis.satelite.util.tipo.Encrypted;

@WebServlet(name = "controller", urlPatterns = {"/controller"})
public final class Controller extends HttpServlet
{
    private static final long serialVersionUID = 1L;
    private static final Logger logger = LogManager.getLogger(Controller.class);
    
    public Controller()
    {
        super();
    }

    @Override
    public void init(ServletConfig config) throws ServletException
    {
        // TODO Auto-generated method stub
    }

    @Override
    public void destroy()
    {
        // TODO Auto-generated method stub
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
    {
        logger.info("doGet[INI] request={}", request );

        ApplicationContext ctx = WebApplicationContextUtils.getWebApplicationContext(request.getServletContext());
        
        ConexionPO po = ctx.getBean(ConexionPO.class);
        ConexionTO to = new ConexionTO();
        
        to.setNombre( "Mi Conexión" );
        to.setUsuario( new Encrypted( "satelite") );
        to.setPassword( new Encrypted( "satelite" ) );
        to.setUrl( new Encrypted("jdbc:oracle:thin:@") );
        to.setDriver( new DriverTO() );
        to.getDriver().setId(1);

        po.guardar(to);
        
        response.getWriter().append("Served at: ").append(request.getContextPath());
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
    {
        doGet(request, response);
    }
}
