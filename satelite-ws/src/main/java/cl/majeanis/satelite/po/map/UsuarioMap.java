package cl.majeanis.satelite.po.map;

import java.util.List;
import java.util.Map;

import cl.majeanis.satelite.to.modelo.UsuarioTO;
import cl.majeanis.satelite.util.po.LegacyMap;

public interface UsuarioMap extends LegacyMap
{
    public List<UsuarioTO> select(Map<String, Object> parm);
}
