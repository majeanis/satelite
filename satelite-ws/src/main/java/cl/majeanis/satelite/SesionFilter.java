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
import javax.servlet.http.HttpServletResponse;

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
        
        /*
         * Nada haremos cuando el request es para el servicio de Sesiones
         */
        String uri = httpRequest.getRequestURI();
        if( uri.matches( ".*/sesiones.*") )
        {
            chain.doFilter(httpRequest, response);
            return;
        }

        /*
         * Para todos los siguientes Request exigiremos un token de Authorization válido
         */
        String authorization = httpRequest.getHeader("Authorization");
        if( authorization == null )
        {
            HttpServletResponse httpResponse = (HttpServletResponse) response;
            httpResponse.setStatus(HttpServletResponse.SC_UNAUTHORIZED);
            return;
        }
        
        Respuesta<SesionTO> sesion = sesionBO.obtener(authorization);
        if( !sesion.isContentOk() )
        {
            HttpServletResponse httpResponse = (HttpServletResponse) response;
            httpResponse.setStatus(HttpServletResponse.SC_FORBIDDEN);
            return;
        }
        
        /*
         * Si llegamos a este punto entonces el token es válido
         * y es posible obtener la sesión del usuario
         */
        httpRequest.setAttribute("sesion", sesion.getContent().get());
        chain.doFilter(httpRequest, response);
    }

    @Override
    public void destroy()
    {
    }
}
