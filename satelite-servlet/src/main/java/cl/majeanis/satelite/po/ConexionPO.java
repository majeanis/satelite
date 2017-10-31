package cl.majeanis.satelite.po;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import cl.majeanis.satelite.po.map.ConexionMap;
import cl.majeanis.satelite.to.modelo.ConexionTO;

@Repository
public class ConexionPO
{
    @Autowired
    private ConexionMap conexion;

    public ConexionTO guardar(ConexionTO data)
    {
        conexion.insert(data);
        return data;
    }
}
