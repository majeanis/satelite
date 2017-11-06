package cl.majeanis.satelite.po;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import cl.majeanis.satelite.po.map.UsuarioMap;
import cl.majeanis.satelite.to.modelo.UsuarioTO;

@Repository
public class UsuarioPO
{
    private static final Logger logger = LogManager.getLogger(UsuarioPO.class);

    @Autowired
    private UsuarioMap usuarioMap;
 
    public UsuarioTO get(String nombre)
    {
        logger.debug("get[INI] nombre={}", nombre );
        
        Map<String, Object> parm = new HashMap<>();
        parm.put("nombre", nombre);
        
        List<UsuarioTO> l = usuarioMap.select(parm);
        if( l.isEmpty() )
        {
            logger.debug("get[FIN] no se encontró registro - nombre={}", nombre);
            return null;
        }
        
        UsuarioTO u = l.get(0);
        logger.debug("get[FIN] registró encontrado={}", u );
        return u;
    }
}
