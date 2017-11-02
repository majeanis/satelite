package cl.majeanis.satelite.util.ws;

import javax.annotation.PostConstruct;
import javax.servlet.ServletContext;
import javax.ws.rs.core.Context;

import org.springframework.context.ApplicationContext;
import org.springframework.web.context.support.WebApplicationContextUtils;

public abstract class RecursoRestBase
{
    @Context
    protected ServletContext servletContext;
    
    protected ApplicationContext appContext;
    
    @PostConstruct
    protected void init()
    {
        if(appContext != null)
            return;
        
        synchronized(RecursoRestBase.class)
        {
            if(appContext != null )
                return;

            appContext = WebApplicationContextUtils.getWebApplicationContext(servletContext);
            if( appContext != null )
                initBeans(appContext);
        }
    }

    abstract protected void initBeans(ApplicationContext appContext);
}
