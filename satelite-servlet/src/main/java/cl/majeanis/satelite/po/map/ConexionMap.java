package cl.majeanis.satelite.po.map;

import java.math.BigInteger;
import java.util.List;

import cl.majeanis.satelite.to.modelo.ConexionTO;
import cl.majeanis.satelite.util.po.SateliteMap;

public interface ConexionMap extends SateliteMap
{
    public void insert(ConexionTO data);
    
    public void update(ConexionTO data);
    
    public ConexionTO select(BigInteger id);
    
    public List<ConexionTO> select();
}
