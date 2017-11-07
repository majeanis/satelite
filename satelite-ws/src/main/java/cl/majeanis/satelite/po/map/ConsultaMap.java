package cl.majeanis.satelite.po.map;

import java.util.List;
import java.util.Map;

import cl.majeanis.satelite.to.modelo.ConsultaTO;
import cl.majeanis.satelite.util.po.LegacyMap;

public interface ConsultaMap extends LegacyMap
{
    public List<ConsultaTO> select(Map<String, Object> parm);
}
