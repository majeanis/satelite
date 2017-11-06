package cl.majeanis.satelite;

import java.io.IOException;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.annotation.WebFilter;
import javax.servlet.http.HttpServletRequest;

import org.springframework.context.ApplicationContext;
import org.springframework.web.context.support.WebApplicationContextUtils;

import cl.majeanis.satelite.bo.SesionBO;
import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.util.Respuesta;

@WebFilter(filterName=".SESION_FILTER.", urlPatterns="/*")
public class SesionFilter implements Filter
{
    private static SesionBO sesionBO;
    
    @Override
    public void init(FilterConfig filterConfig) throws ServletException
    {
        if( sesionBO != null )
            return;

        synchronized(SesionFilter.class)
        {
            if( sesionBO != null )
                return;

            ServletContext context = filterConfig.getServletContext();
            ApplicationContext appCtx = WebApplicationContextUtils.getWebApplicationContext(context);
            sesionBO = appCtx.getBean(SesionBO.class);
        }
    }

    @Override
    public void doFilter(ServletRequest request, 
                         ServletResponse response, 
                         FilterChain chain)
            throws IOException, ServletException
    {
        HttpServletRequest httpRequest = (HttpServletRequest) request;
        String authorization = httpRequest.getHeader("Authorization");

        Respuesta<SesionTO> sesion = sesionBO.obtener(authorization);
        if( sesion.isContentOk() )
        {
            httpRequest.setAttribute("sesion", sesion.getContent().get());
        }

        chain.doFilter(httpRequest, response);
    }

    @Override
    public void destroy()
    {
    }
}
