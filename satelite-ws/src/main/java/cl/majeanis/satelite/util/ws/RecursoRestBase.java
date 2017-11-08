package cl.majeanis.satelite.util.ws;

import javax.annotation.PostConstruct;
import javax.servlet.http.HttpServletRequest;
import javax.ws.rs.core.Context;

import org.springframework.context.ApplicationContext;
import org.springframework.web.context.support.WebApplicationContextUtils;

import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;

public abstract class RecursoRestBase
{
    @Context
    protected HttpServletRequest servletRequest;
    
    protected ApplicationContext appContext;
    
    protected SesionTO sesionRequest;
    
    @PostConstruct
    protected void init()
    {
        appContext = WebApplicationContextUtils.getWebApplicationContext(servletRequest.getServletContext());
        sesionRequest = (SesionTO) servletRequest.getAttribute("sesion");
        initBeans(appContext);
    }

    protected Resultado checkSesion()
    {
        Resultado r = new ResultadoProceso();
        if ( sesionRequest == null )
        {
            r.addError( "No ha informado datos de la sesión" );
            return r;
        }

        return r;
    }

    abstract protected void initBeans(ApplicationContext appContext);
}
