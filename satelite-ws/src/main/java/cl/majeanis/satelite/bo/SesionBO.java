package cl.majeanis.satelite.bo;

import javax.naming.AuthenticationException;
import javax.naming.CommunicationException;
import javax.naming.NamingException;
import javax.naming.ldap.LdapContext;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import cl.majeanis.satelite.po.SesionPO;
import cl.majeanis.satelite.po.UsuarioPO;
import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.to.modelo.UsuarioTO;
import cl.majeanis.satelite.util.LoginUtils;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;
import cl.majeanis.satelite.util.SisProperties;

@Service
public class SesionBO
{
    private static final Logger logger = LogManager.getLogger(SesionBO.class);
    
    @Autowired
    private SesionPO sesionPO;
    
    @Autowired
    private UsuarioPO usuarioPO;
    
    @Autowired
    private SisProperties properties;
    
    public Respuesta<SesionTO> autenticar(String authorization)
    {
        logger.info("autenticar[INI] authorization={}", authorization );

        Resultado rtdo = new ResultadoProceso();
        if( authorization == null || StringUtils.isEmpty(authorization) )
        {
            rtdo.addError("Debe informar datos para la autenticación");
            logger.info("autenticar[FIN] authorization vacío" );
            return new Respuesta<>(rtdo);
        }
        
        String serverAd = properties.getServerAD();
        String domainAd = properties.getDomainAD();
        logger.debug("autenticar: serverAd={} domainAd={}", serverAd, domainAd );

        try
        {
            String[] credencial = authorization.split(":");
            logger.trace("autenticar: credenciales={}", (Object[]) credencial );
            if( credencial != null & credencial.length != 2)
            {
                rtdo.addError("Credenciales informadas no son correctas");
                logger.info("autenticar[FIN] split del token authorization generó arreglo incorrecto - credencial={} authorization={}", (Object[]) credencial, authorization);
                return new Respuesta<>(rtdo);
            }
            
            /**
             * Procedemos a validar las credenciales de autenticación
             */
            LdapContext ctx = LoginUtils.loginToAd(credencial[0], credencial[1], domainAd, serverAd);
            ctx.close();
            ctx = null;
            
            /**
             * Si llegamos a este punto, entonces las credenciales son correctas
             * y tenemos que buscar al usuario en la BD. Si el usuario no existe
             * en la BD entonces no es factible autenticarlo
             */
            UsuarioTO usua = usuarioPO.get(credencial[0]);
            if( usua == null )
            {
                rtdo.addError("Usuario \"%1$s\" no existe en la aplicación",  credencial[0]);
                logger.info("autenticar[FIN] usuario no existe en la BD - usuario={}", credencial[0] );
            }
            
            SesionTO sesion = sesionPO.crear(usua);
            logger.info("autenticar[FIN] sesion={}", sesion);

            return new Respuesta<>(sesion);
        } catch (CommunicationException e)
        {
            rtdo.addError("No hay comunicación con el servidor Active Directory [%1$s:%2$s]", serverAd, domainAd);
            logger.error("autenticar[ER2] al autenticar usuario - " + rtdo, e);

        } catch (AuthenticationException e)
        {
            rtdo.addError("No fue posibe autenticar al usuario");
            logger.error("autenticar[ER3] al autenticar usuario - " + rtdo, e);

        } catch (NamingException e)
        {
            rtdo.addError("Error al autenticar al usuario");
            logger.error("autenticar[ER4] al autenticar usuario - " + rtdo, e);
        }
        
        return new Respuesta<>(rtdo);
    }
    
    public Respuesta<SesionTO> obtener(String id)
    {
        logger.info("obtener[INI] id={}", id );
        
        Resultado rtdo = new ResultadoProceso();

        SesionTO sesion = sesionPO.get(id);
        if( sesion == null )
        {
            rtdo.addError("No existe sesión con id=%1$s", id);
            logger.info("obtener[FIN] no existe sesión con id={}", id );
            return new Respuesta<>(rtdo);
        }

        logger.info("obtener[FIN] sesion={}", sesion );
        return new Respuesta<>(sesion);
    }
}
