package cl.majeanis.satelite.po;

import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import cl.majeanis.satelite.po.map.ConexionMap;
import cl.majeanis.satelite.to.modelo.ConexionTO;
import cl.majeanis.satelite.util.Utils;

@Repository
public class ConexionPO
{
    private static final Logger logger = LogManager.getLogger(ConexionPO.class);

    @Autowired
    private ConexionMap conxMap;

    public ConexionTO guardar(ConexionTO data)
    {
        logger.debug("guardar[INI] data={}", data );
        
        if( data.getId() == null )
        {
            conxMap.insert(data);
        } else
        {
            conxMap.update(data);
        }

        logger.debug("guardar[FIN] data={}", data );
        return data;
    }
    
    public ConexionTO get(BigInteger id)
    {
        logger.debug("get[INI] id={}", id );
        
        Map<String, Object> parm = new HashMap<>();
        parm.put("id", id);
     
        List<ConexionTO> l = conxMap.select(parm);
        logger.debug("get[FIN] registros={}", Utils.sizeOf(l) );

        if( l.isEmpty() )
            return null;
        
        return l.get(0);
    }
    
    public List<ConexionTO> getList()
    {
        logger.debug("getList[INI]");
        
        Map<String, Object> parm = new HashMap<>();
        List<ConexionTO> l = conxMap.select(parm);
        
        logger.debug("getList[FIN] registros={}", Utils.sizeOf(l) );
        return l;
    }
}
