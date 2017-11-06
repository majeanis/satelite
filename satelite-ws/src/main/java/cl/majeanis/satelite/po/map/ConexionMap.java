package cl.majeanis.satelite.po.map;

import java.util.List;
import java.util.Map;

import cl.majeanis.satelite.to.modelo.ConexionTO;
import cl.majeanis.satelite.util.po.SateliteMap;

public interface ConexionMap extends SateliteMap
{
    public void insert(ConexionTO data);
    
    public void update(ConexionTO data);
    
    public List<ConexionTO> select(Map<String, Object> parm);
}
