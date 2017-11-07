package cl.majeanis.satelite.po;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    public List<ConsultaTO> getList(String baseDatosId, String usuario)
    {
        logger.debug("getList[INI] usuario={}", usuario);

        Map<String, Object> parm = new HashMap<>();
        parm.put("usuario", usuario);
        parm.put("baseDatosId", baseDatosId);
        
        List<ConsultaTO> l = consMap.select(parm);
        
        logger.debug("getList[FIN] registros retornados={}", Utils.sizeOf(l));
        return l;
    }
    
    public List<ConsultaTO> getList(String baseDatosId)
    {
        return getList(baseDatosId, null);
    }
}
