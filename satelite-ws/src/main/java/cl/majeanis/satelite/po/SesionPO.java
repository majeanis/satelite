package cl.majeanis.satelite.po;

import java.time.LocalDateTime;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.stereotype.Repository;

import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.to.modelo.UsuarioTO;
import cl.majeanis.satelite.util.tipo.Encrypted;

@Repository
public class SesionPO
{
    private static final Logger logger = LogManager.getLogger(SesionPO.class);

    private static final ConcurrentMap<String, SesionTO> cacheSesion = new ConcurrentHashMap<>();
    
    public SesionTO crear(UsuarioTO usuario)
    {
        logger.debug("crear[INI] usuario={}", usuario);
        
        String plainToken = usuario.getNombre() + ":" + usuario.getTipo().getCodigo();
        logger.trace("crear: despues de crear el plainToken={}", plainToken );

        SesionTO sesion = new SesionTO();
        sesion.setId( new Encrypted( plainToken ).text() );
        sesion.setUsuario(usuario);
        sesion.setFecha(LocalDateTime.now());

        cacheSesion.putIfAbsent( sesion.getId(), sesion);

        logger.debug("crear[FIN] sesion={}", sesion );
        return sesion;        
    }
    
    public SesionTO get(String id)
    {
        logger.debug("get[INI] id={}", id );
        
        if( id == null )
            return null;

        SesionTO sesion = cacheSesion.get(id);

        logger.debug("get[FIN] sesion={}",  sesion );
        return sesion;
    }
}
