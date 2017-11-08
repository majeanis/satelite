package cl.majeanis.satelite.po;

import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import cl.majeanis.satelite.po.map.ConsultaMap;
import cl.majeanis.satelite.to.modelo.ConsultaTO;
import cl.majeanis.satelite.util.Utils;

@Repository
public class ConsultaPO
{
    private static final Logger logger = LogManager.getLogger(ConsultaPO.class);

    @Autowired
    private ConsultaMap consMap;

    public List<ConsultaTO> getList(Optional<String> usuario, Optional<String> baseDatosId)
    {
        logger.debug("getList[INI] usuario={} baseDatosId={}", usuario, baseDatosId);

        Map<String, Object> parm = new HashMap<>();
        parm.put("usuario", usuario.orElse(null));
        parm.put("baseDatosId", baseDatosId.orElse(null));
        
        List<ConsultaTO> l = consMap.select(parm);
        
        logger.debug("getList[FIN] registros retornados={}", Utils.sizeOf(l));
        return l;
    }
    
    public ConsultaTO get(BigInteger consultaId, Optional<String> usuario)
    {
        logger.debug("get[INI] consultaId={} usuario={}", consultaId, usuario );
        
        Map<String, Object> parm = new HashMap<>();
        parm.put("consultaId", consultaId);
        parm.put("usuario", usuario.orElse(null));

        List<ConsultaTO> l = consMap.select(parm);
        if( l.isEmpty() )
        {
            logger.debug("get[FIN] no existe registro de la consulta - consultaId={}", consultaId );
            return null;
        }
        
        ConsultaTO o = l.get(0);
        logger.debug("get[FIN] registro encontrado {}", o );
        return o;
    }
}
