package cl.majeanis.satelite.po;

import java.time.LocalDateTime;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.stereotype.Repository;

import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.to.modelo.UsuarioTO;
import cl.majeanis.satelite.util.JsonUtils;
import cl.majeanis.satelite.util.tipo.Encrypted;

@Repository
public class SesionPO
{
    private static final Logger logger = LogManager.getLogger(SesionPO.class);

    public SesionTO crear(UsuarioTO usuario)
    {
        logger.debug("crear[INI] usuario={}", usuario);
        
        SesionTO sesion = new SesionTO();
        sesion.setUsuario(usuario);
        sesion.setFecha(LocalDateTime.now());
        
        Encrypted token = new Encrypted( JsonUtils.toJson(usuario) );
        sesion.setId( token.text() );

        logger.debug("crear[FIN] sesion={}", sesion );
        return sesion;        
    }
}
