package cl.majeanis.satelite.bo;

import java.io.UnsupportedEncodingException;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import cl.majeanis.satelite.po.SesionPO;
import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.to.modelo.TipoUsuarioTO;
import cl.majeanis.satelite.to.modelo.UsuarioTO;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;

@Service
public class SesionBO
{
    private static final Logger logger = LogManager.getLogger(SesionBO.class);
    
    @Autowired
    private SesionPO sesionPO;
    
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
        
        try
        {
            String decoded = new String( Base64.decodeBase64(authorization), "UTF-8");
            logger.trace("autenticar: despues de Base64 decode={}", decoded );
            String[] credencial = decoded.split(":");
            
            logger.trace("autenticar: credenciales={}", (Object[]) credencial );
            
            TipoUsuarioTO tu = new TipoUsuarioTO();
            tu.setNombre("ADMIN");

            UsuarioTO u = new UsuarioTO();
            u.setNombre("mauricio.camara");
            u.setTipo(tu);
            
            SesionTO sesion = sesionPO.crear(u);
            logger.info("autenticar[FIN] sesion={}", sesion);

            return new Respuesta<>(sesion);
        } catch (UnsupportedEncodingException e)
        {
            logger.error("autenticar[ERR]", e);
            rtdo.addError("No fue posibe decodicar la llave Authorization");

            return new Respuesta<>(rtdo);
        }
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
